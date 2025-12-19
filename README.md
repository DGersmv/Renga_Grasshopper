# Renga_Grasshopper Integration

Интеграция между Grasshopper и Renga для автоматического размещения колонн в Renga на основе точек, заданных в Grasshopper.

## Описание проекта

Проект состоит из двух плагинов:
- **Плагин Renga** - TCP сервер, который принимает данные от Grasshopper и создает/обновляет колонны
- **Плагин Grasshopper** - два компонента для подключения к Renga и отправки данных о колоннах

## Структура проекта

```
Renga_Grasshopper/
├── Renga/              # Плагин для Renga
│   ├── RengaPlugin.cs
│   ├── Renga.csproj
│   └── Renga.rndesc
├── Grasshopper/        # Плагин для Grasshopper
│   ├── Components/
│   │   ├── RengaConnectComponent.cs
│   │   └── RengaCreateColumnsComponent.cs
│   ├── Client/
│   │   └── RengaGhClient.cs
│   ├── Properties/
│   │   ├── AssemblyInfo.cs
│   │   └── GrasshopperInfo.cs
│   └── Grasshopper.csproj
└── Renga_Grasshopper.sln
```

## Требования

- **Renga SDK** - установлен в стандартном месте
- **Rhino 8** с Grasshopper
- **Visual Studio 2022** или новее
- **.NET 8.0** (для плагина Renga)
- **.NET Framework 4.8** (для плагина Grasshopper)

## Сборка проекта

1. Откройте `Renga_Grasshopper.sln` в Visual Studio
2. Убедитесь, что пути к библиотекам Grasshopper и RhinoCommon корректны в `Grasshopper.csproj`
3. Соберите решение (Build Solution)

### Установка плагина Renga

После сборки:
1. Скопируйте `Renga.dll` и `Renga.rndesc` из `bin\Release\Renga\` в папку плагинов Renga
2. Перезапустите Renga

### Установка плагина Grasshopper

После сборки плагин автоматически копируется в `%APPDATA%\Grasshopper\Libraries\` (благодаря PostBuildEvent).
Или скопируйте `Renga_Grasshopper.gha` вручную.

## Использование

### В Renga:
1. Плагин автоматически запускает TCP сервер на порту 50100 при загрузке
2. (TODO: Добавить UI для настройки порта и управления сервером)

### В Grasshopper:
1. Добавьте компонент **"Renga Connect"**:
   - Укажите порт (по умолчанию 50100)
   - Включите подключение (Connect = true)
2. Добавьте компонент **"Renga Create Columns"**:
   - Подключите точки (Points)
   - Подключите выход Client от компонента "Renga Connect"
   - Компонент автоматически отправит данные в Renga

## Функциональность

- ✅ Создание колонн в Renga на основе точек из Grasshopper
- ✅ Отслеживание точек по GUID для обновления позиций
- ✅ Автоматическое обновление при изменении точек
- ✅ Не создает дубликаты колонн для существующих точек
- ⏳ UI для настройки порта в Renga (в разработке)
- ⏳ Кнопки Update в компонентах Grasshopper (в разработке)

## Технические детали

- **Протокол**: TCP/IP
- **Порт по умолчанию**: 50100
- **Формат данных**: JSON
- **Отслеживание**: GUID точек для синхронизации колонн

## Лицензия

Copyright © 2025 Renga Software LLC. All rights reserved.










