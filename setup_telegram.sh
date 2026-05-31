#!/bin/bash
set -e

if [ $# -lt 2 ]; then
    echo "Использование: $0 <BOT_TOKEN> <CHAT_ID>"
    echo ""
    echo "Пример:"
    echo "  $0 123456:ABC-DEF1234ghIkl-zyx57W2v1u123ew11 123456789"
    echo ""
    echo "Где взять:"
    echo "  BOT_TOKEN — создать бота у @BotFather в Telegram"
    echo "  CHAT_ID  — написать боту, открыть https://api.telegram.org/botТОКЕН/getUpdates"
    exit 1
fi

BOT_TOKEN=$1
CHAT_ID=$2

# 1. Добавляем в ~/.bashrc
echo '' >> ~/.bashrc
echo "# Telegram notifications for site parser" >> ~/.bashrc
echo "export TELEGRAM_BOT_TOKEN=\"$BOT_TOKEN\"" >> ~/.bashrc
echo "export TELEGRAM_CHAT_ID=\"$CHAT_ID\"" >> ~/.bashrc
echo "Добавлено в ~/.bashrc"

# 2. Обновляем parser.service
SERVICE_FILE="/root/parser/parser.service"
if [ -f "$SERVICE_FILE" ]; then
    sed -i "s/Environment=TELEGRAM_BOT_TOKEN=.*/Environment=TELEGRAM_BOT_TOKEN=$BOT_TOKEN/" "$SERVICE_FILE"
    sed -i "s/Environment=TELEGRAM_CHAT_ID=.*/Environment=TELEGRAM_CHAT_ID=$CHAT_ID/" "$SERVICE_FILE"
    echo "Обновлён $SERVICE_FILE"
    echo "Перезапустите сервис: sudo systemctl daemon-reload && sudo systemctl restart parser"
else
    echo "Файл $SERVICE_FILE не найден. Обновите вручную."
fi

echo ""
echo "Готово! Telegram-уведомления настроены."
