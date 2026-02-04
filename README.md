# EpidMonitor Updater

Автообновляющийся лаунчер для EpidMonitor.

## Как работает
- Лаунчер читает manifest.json из GitHub Releases
- Скачивает zip релиза
- Проверяет SHA256
- Распаковывает в AppData
- Запускает приложение

Git используется только для кода лаунчера.
