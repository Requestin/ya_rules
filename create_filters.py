#!/usr/bin/env python3
"""Создаёт правила Яндекс Почты по таблице tablica.xlsx через открытый браузер.

Вход в аккаунт скрипт не выполняет: в открывшемся окне нужно войти вручную.
Дальше для каждой строки таблицы создаётся правило
«От кого содержит адрес → Положить в папку».
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

import openpyxl
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent
DEFAULT_XLSX = ROOT / "tablica.xlsx"
PROFILE_DIR = ROOT / ".mail-profile"
DEBUG_DIR = ROOT / "debug"
MAIL_URL = "https://mail.yandex.ru/"
FILTERS_URL = "https://mail.yandex.ru/#setup/filters"
ACCOUNT = "k.darchinyants@gyhyry.com"


def load_rules(path: Path) -> list[tuple[str, str]]:
    workbook = openpyxl.load_workbook(path, data_only=True)
    sheet = workbook.active
    rules: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        folder = str(row[0]).strip() if row and row[0] else ""
        email = str(row[1]).strip() if row and len(row) > 1 and row[1] else ""
        if not folder and not email:
            continue
        if not folder or "@" not in email:
            raise SystemExit(f"Строка {row_number}: нужны папка и адрес почты, сейчас {row!r}")
        key = (folder, email.lower())
        if key in seen:
            continue
        seen.add(key)
        rules.append((folder, email))
    if not rules:
        raise SystemExit(f"В {path} нет строк с папкой и адресом")
    return rules


def click_named(scope, name: str) -> None:
    for role in ("button", "link"):
        locator = scope.get_by_role(role, name=name, exact=True)
        if locator.count() and locator.first.is_visible():
            locator.first.click()
            return
    scope.get_by_text(name, exact=True).first.click()


def dismiss_dialogs(page) -> None:
    for text in ("Позже", "Не сейчас", "Пропустить"):
        dialog = page.get_by_role("dialog")
        if not dialog.count() or not dialog.first.is_visible():
            return
        button = dialog.first.get_by_role("button", name=text, exact=True)
        if button.count() and button.first.is_visible():
            button.first.click()
            page.wait_for_timeout(300)


def body_text(scope) -> str:
    try:
        return scope.locator("body").inner_text(timeout=5000)
    except PlaywrightTimeout:
        return ""


def scope_with_text(page, text: str):
    if text in body_text(page):
        return page
    for frame in page.frames:
        try:
            if text in (frame.locator("body").inner_text(timeout=1000) or ""):
                return frame
        except Exception:
            continue
    return page


def wait_until_logged_in(page) -> None:
    print()
    print(f"Войдите в ящик {ACCOUNT}, если окно ещё на странице входа.")
    print("Капчу и код подтверждения вводите сами. Жду почту до 10 минут.")
    print()
    deadline = time.time() + 600
    while time.time() < deadline:
        dismiss_dialogs(page)
        current = page.url
        text = body_text(page)
        on_mail = "mail.yandex." in current and "passport.yandex." not in current
        if on_mail and ("Входящие" in text or "Написать" in text or "Правила обработки писем" in text):
            print(f"Почта открыта: {current}")
            return
        page.wait_for_timeout(1000)
    raise SystemExit("Не дождался входа в почту. Запустите скрипт ещё раз и войдите в открывшемся окне.")


def open_filters(page):
    page.goto(FILTERS_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)
    dismiss_dialogs(page)
    scope = scope_with_text(page, "Создать правило")
    if "Создать правило" in body_text(scope) or "Правила обработки писем" in body_text(scope):
        return scope

    for label in ("Настройки", "Все настройки"):
        try:
            click_named(page, label)
            page.wait_for_timeout(700)
        except Exception:
            continue
    try:
        click_named(page, "Правила обработки писем")
        page.wait_for_timeout(1000)
    except Exception:
        pass
    scope = scope_with_text(page, "Создать правило")
    if "Создать правило" not in body_text(scope):
        save_debug(page, "filters-not-found")
        raise SystemExit(
            "Не открылась страница «Правила обработки писем». "
            f"Снимок экрана: {DEBUG_DIR / 'filters-not-found.png'}"
        )
    return scope


def save_debug(page, name: str) -> None:
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{name}.png"), full_page=True)
    except Exception:
        pass
    try:
        (DEBUG_DIR / f"{name}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass
    print(f"Снимок для разбора: {DEBUG_DIR / (name + '.png')}")


def visible_selects(scope):
    found = []
    locator = scope.locator("select")
    for index in range(locator.count()):
        item = locator.nth(index)
        if item.is_visible():
            found.append(item)
    return found


def option_labels(select) -> list[str]:
    return [text.strip() for text in select.locator("option").all_inner_texts()]


def choose_in_selects(scope, label: str) -> bool:
    for select in visible_selects(scope):
        labels = option_labels(select)
        if label in labels:
            select.select_option(label=label)
            return True
    return False


def owner_page(scope):
    return scope.page if not hasattr(scope, "frames") else scope


def choose_custom(scope, label: str) -> bool:
    """Выбирает пункт уже открытого списка."""
    option = scope.get_by_role("option", name=label, exact=True)
    if option.count() and option.first.is_visible():
        option.first.click()
        return True
    menu_item = scope.get_by_role("menuitem", name=label, exact=True)
    if menu_item.count() and menu_item.first.is_visible():
        menu_item.first.click()
        return True
    popup = owner_page(scope).locator("[role='listbox'], [role='menu']")
    if popup.count():
        item = popup.last.get_by_text(label, exact=True)
        if item.count() and item.first.is_visible():
            item.first.click()
            return True
    return False


def choice_already_selected(scope, label: str) -> bool:
    for select in visible_selects(scope):
        selected = select.evaluate(
            "el => (el.options[el.selectedIndex] && el.options[el.selectedIndex].text || '').trim()"
        )
        if selected == label:
            return True
    button = scope.get_by_role("button", name=label, exact=True)
    return bool(button.count() and button.first.is_visible())


def ensure_choice(scope, label: str, openers: tuple[str, ...]) -> None:
    if choice_already_selected(scope, label) or choose_in_selects(scope, label):
        return
    page = owner_page(scope)
    for opener in openers:
        control = scope.get_by_role("button", name=opener, exact=True)
        if not control.count() or not control.first.is_visible():
            control = scope.get_by_text(opener, exact=True)
        if not control.count() or not control.first.is_visible():
            continue
        control.first.click()
        page.wait_for_timeout(300)
        if choose_custom(page, label) or choose_in_selects(scope, label):
            return
    raise RuntimeError(f"Не удалось выбрать «{label}»")


def mark_form(page) -> bool:
    script = """() => {
      document.querySelectorAll('[data-ya-filter-form]').forEach(el => el.removeAttribute('data-ya-filter-form'));
      const needles = ['Положить в папку', 'Если'];
      let best = null;
      let bestLen = Infinity;
      for (const el of document.querySelectorAll('div, form, section, main')) {
        const text = el.innerText || '';
        if (text.length < 40 || text.length >= bestLen) continue;
        if (!needles.every(needle => text.includes(needle))) continue;
        best = el;
        bestLen = text.length;
      }
      if (!best) return false;
      best.setAttribute('data-ya-filter-form', '1');
      return true;
    }"""
    for frame in page.frames:
        try:
            if frame.evaluate(script):
                return True
        except Exception:
            continue
    return False


def form_locator(page):
    for frame in page.frames:
        locator = frame.locator("[data-ya-filter-form='1']")
        if locator.count():
            return locator.first
    raise RuntimeError("Форма правила не найдена")


def topmost_text_input(scope):
    locator = scope.locator("input")
    best = None
    best_y = 10**9
    for index in range(locator.count()):
        item = locator.nth(index)
        if not item.is_visible():
            continue
        input_type = (item.get_attribute("type") or "text").lower()
        if input_type not in ("text", "search", "email"):
            continue
        box = item.bounding_box()
        if not box:
            continue
        if box["y"] < best_y:
            best_y = box["y"]
            best = item
    if best is None:
        raise RuntimeError("Не нашёл поле, куда вводить адрес отправителя")
    return best


def set_from_contains(scope, email: str) -> None:
    ensure_choice(scope, "От кого", ("Тема", "Кому", "Копия", "Заголовок", "Тело письма"))
    ensure_choice(scope, "содержит", ("совпадает с", "не содержит", "не совпадает с", "начинается с"))
    field = topmost_text_input(scope)
    field.click()
    field.fill(email)


def enable_move_to_folder(scope, folder: str) -> None:
    page = owner_page(scope)
    label = scope.locator("label").filter(has_text=re.compile(r"Положить в папку"))
    if label.count():
        checkbox = label.first.locator("input[type=checkbox]")
        if checkbox.count():
            checkbox.first.check(force=True)
        else:
            label.first.click()
    else:
        scope.get_by_text("Положить в папку", exact=True).first.click()
    page.wait_for_timeout(400)

    if choose_in_selects(scope, folder):
        return

    # Список папок у Яндекса часто рисуется отдельным слоем поверх страницы.
    for select in visible_selects(page):
        labels = option_labels(select)
        if folder in labels:
            select.select_option(label=folder)
            return
    picker = scope.get_by_text(re.compile(r"Выберите папку|Входящие"))
    if picker.count() and picker.first.is_visible():
        picker.first.click()
    else:
        move = scope.get_by_text("Положить в папку", exact=True).first
        sibling = move.locator("xpath=following::*[@role='button' or self::button][1]")
        if sibling.count():
            sibling.click()
    page.wait_for_timeout(300)
    option = page.get_by_role("option", name=folder, exact=True)
    if option.count() and option.first.is_visible():
        option.first.click()
        return
    text = page.get_by_text(folder, exact=True)
    if text.count():
        text.last.click()
        return
    raise RuntimeError(f"В списке нет папки «{folder}». Создайте её в почте и запустите скрипт снова.")


def fill_rule_name(scope, name: str) -> None:
    labeled = scope.get_by_label(re.compile(r"Название"))
    if labeled.count() and labeled.first.is_visible():
        labeled.first.fill(name)
        return
    inputs = scope.locator("input")
    lowest = None
    lowest_y = -1
    for index in range(inputs.count()):
        item = inputs.nth(index)
        if not item.is_visible():
            continue
        input_type = (item.get_attribute("type") or "text").lower()
        if input_type not in ("text", "search", "email"):
            continue
        box = item.bounding_box()
        if box and box["y"] > lowest_y:
            lowest_y = box["y"]
            lowest = item
    if lowest is not None and (lowest.input_value() or "") == "":
        lowest.fill(name)


def create_rule(page, scope, folder: str, email: str, apply_existing: bool) -> None:
    rule_name = f"{folder}: {email}"
    print(f"Создаю правило: от «{email}» → папка «{folder}»")
    click_named(scope, "Создать правило")
    page.wait_for_timeout(800)
    if not mark_form(page):
        save_debug(page, "form-missing")
        raise RuntimeError("После «Создать правило» не появилась форма с действием «Положить в папку»")
    form = form_locator(page)
    save_debug(page, "form")
    set_from_contains(form, email)
    enable_move_to_folder(form, folder)
    if apply_existing:
        existing = form.get_by_text("Применить к существующим письмам", exact=False)
        if existing.count():
            existing.first.click()
    fill_rule_name(form, rule_name)
    save_debug(page, "before-submit")
    buttons = form.get_by_role("button", name="Создать правило", exact=True)
    if buttons.count():
        buttons.last.click()
    else:
        click_named(form, "Создать правило")
    page.wait_for_timeout(1500)
    text = body_text(scope_with_text(page, "Правила обработки писем"))
    if rule_name not in text and email not in text:
        save_debug(page, "after-submit")
        raise RuntimeError(
            "После нажатия «Создать правило» адрес не появился в списке. "
            "Смотрите снимок debug/after-submit.png — форму, скорее всего, не приняли."
        )
    print(f"Готово: {rule_name}")


def already_present(scope, email: str) -> bool:
    return email.lower() in body_text(scope).lower()


def main() -> None:
    parser = argparse.ArgumentParser(description="Создать правила Яндекс Почты по tablica.xlsx")
    parser.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX)
    parser.add_argument(
        "--apply-existing",
        action="store_true",
        help="включить «Применить к существующим письмам»",
    )
    parser.add_argument("--dry-run", action="store_true", help="только показать, какие правила будут созданы")
    args = parser.parse_args()

    rules = load_rules(args.xlsx)
    print(f"В таблице {len(rules)} правил для {ACCOUNT}:")
    for folder, email in rules:
        print(f"  {email}  →  {folder}")
    if args.dry_run:
        return

    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            locale="ru-RU",
            viewport={"width": 1440, "height": 960},
            slow_mo=80,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(15000)
        try:
            page.goto(MAIL_URL, wait_until="domcontentloaded")
            wait_until_logged_in(page)
            created = 0
            for folder, email in rules:
                scope = open_filters(page)
                if already_present(scope, email):
                    print(f"Пропускаю {email}: такой адрес уже есть на странице правил")
                    continue
                create_rule(page, scope, folder, email, args.apply_existing)
                created += 1
            print()
            print(f"Создано правил: {created}. Окно закроется через 20 секунд.")
            page.wait_for_timeout(20000)
        except Exception as error:
            save_debug(page, "error")
            print(f"Ошибка: {error}", file=sys.stderr)
            print("Окно останется открытым 2 минуты, чтобы можно было посмотреть, на чём остановилось.")
            try:
                page.wait_for_timeout(120000)
            except Exception:
                pass
            raise
        finally:
            context.close()


if __name__ == "__main__":
    main()
