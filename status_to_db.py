# -*- coding: utf-8 -*-
"""
Мониторинг статуса пользователя с сохранением в базу данных.
Запускает периодическую проверку статуса для указанного URL.
"""
import argparse
import sys
import time
from datetime import datetime
from zoneinfo import ZoneInfo

TZ_EKB = ZoneInfo("Asia/Yekaterinburg")


def now_ekb() -> str:
    return datetime.now(TZ_EKB).strftime('%Y-%m-%d %H:%M:%S')

from site_parser import parse_user_status, save_status_to_db, get_history_summary_from_db, export_to_excel

ONLINE_TAGS = [
    {'selector': 'div.status-online', 'text': 'Онлайн'},
    {'selector': 'span.online-indicator', 'text': 'active'},
    {'selector': 'div.user-status', 'text': 'Жду звонка'},
    {'selector': 'div.presence', 'text': 'Доступен'},
    {'selector': 'div.activity-status', 'text': 'Только что'},
]


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')

    parser = argparse.ArgumentParser(description="Мониторинг статуса пользователя с сохранением в БД")
    parser.add_argument('url', nargs='?', help='URL профиля для мониторинга')
    parser.add_argument('--interval', type=int, default=300, help='Интервал проверки в секундах (по умолчанию 300)')
    parser.add_argument('--count', type=int, default=0, help='Количество проверок (0 = бесконечно)')
    parser.add_argument('--export', nargs='?', const='all', metavar='URL', help='Экспорт данных в Excel (без аргумента — все URL)')
    parser.add_argument('--output', default='user_status_export.xlsx', help='Имя файла для экспорта')

    args = parser.parse_args()

    if args.export is not None:
        url_filter = None if args.export == 'all' else args.export
        print(f"Экспорт данных в Excel: {args.output}" + (f" (фильтр: {url_filter})" if url_filter else " (все URL)"))
        excel_bytes = export_to_excel(url=url_filter)
        with open(args.output, 'wb') as f:
            f.write(excel_bytes)
        print(f"Готово: {args.output}")
        return

    print(f"Мониторинг: {args.url}")
    print(f"Интервал: {args.interval} сек")
    print("Нажмите Ctrl+C для остановки\n")

    checks = 0
    try:
        while True:
            checks += 1
            print(f"[{now_ekb()}] Проверка #{checks}...")

            result = parse_user_status(args.url, ONLINE_TAGS)
            save_status_to_db(result)

            if result['success']:
                status_ru = "ONLINE" if result['status'] == 'online' else "OFFLINE"
                print(f"  -> Статус: {status_ru}")
                if result['online_indicators']:
                    for ind in result['online_indicators']:
                        print(f"     {ind['selector']}: \"{ind.get('text', ind.get('content', ''))}\"")
            else:
                print(f"  -> Ошибка: {result.get('error', 'Неизвестная ошибка')}")

            if args.count and checks >= args.count:
                break

            time.sleep(args.interval)

    except KeyboardInterrupt:
        print("\n\nМониторинг остановлен.")

    summary = get_history_summary_from_db(args.url)
    if summary['total'] > 0:
        print(f"\nИтого проверок: {summary['total']}")
        print(f"Онлайн: {summary['online']}, Оффлайн: {summary['offline']}, Неизвестно: {summary['unknown']}")


if __name__ == '__main__':
    main()
