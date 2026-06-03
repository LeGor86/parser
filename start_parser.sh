#!/bin/bash

# Автоматический запуск парсера (Linux/Ubuntu)
export TZ=Asia/Yekaterinburg

# Аутентификация (задайте свои логин и пароль)
# export AUTH_USERNAME=admin
# export AUTH_PASSWORD=your_password

# Установка зависимостей
echo "Установка необходимых зависимостей..."
pip3 install requests beautifulsoup4 openpyxl

if [ $? -ne 0 ]; then
    echo "Ошибка установки зависимостей"
    read -p "Нажмите Enter для выхода..."
    exit 1
fi

echo ""
echo "Выберите режим запуска:"
echo "1 - Веб-интерфейс (сервер)"
echo "2 - Мониторинг одного URL в консоли"
read -p "Введите 1 или 2: " MODE

if [ "$MODE" = "2" ]; then
    read -p "Введите URL профиля для мониторинга: " URL
    if [ -n "$URL" ]; then
        read -p "Интервал проверки в секундах (по умолчанию 300): " INTERVAL
        INTERVAL=${INTERVAL:-300}
        echo "Запуск мониторинга..."
        python3 status_to_db.py "$URL" --interval "$INTERVAL"
    fi
else
    echo "Запуск веб-сервера..."
    python3 site_parser.py --serve
fi

read -p "Нажмите Enter для выхода..."
