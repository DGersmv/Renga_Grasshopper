# Новая архитектура плагина Renga-Grasshopper

## Проблемы текущей архитектуры

1. **Проблема с TCP коммуникацией**: Клиент не получает ответы от сервера (timeout)
2. **Смешанная ответственность**: Все функции в одном файле
3. **Сложность отладки**: Нет четкого разделения между модулями
4. **Проблемы с синхронизацией**: Неправильная обработка чтения/записи в TCP

## Новая архитектура

### Принципы

1. **Разделение ответственности**: Каждый модуль отвечает за одну функцию
2. **Четкий протокол**: Стандартизированный формат сообщений
3. **Надежная связь**: Правильная обработка TCP соединений
4. **Модульность**: Легко добавлять новые функции

---

## Структура модулей

### 1. Модуль связи (Connection Module)

**Назначение**: Управление TCP соединением между Grasshopper и Renga

**Компоненты**:
- `RengaConnectionClient` - TCP клиент для Grasshopper
- `RengaConnectionServer` - TCP сервер для Renga
- `ConnectionProtocol` - протокол обмена данными

**Функции**:
- Установка соединения
- Отправка/получение сообщений
- Обработка ошибок соединения
- Переподключение при разрыве

**Протокол**:
```
Request Format:
{
  "id": "unique-request-id",
  "command": "command-name",
  "data": { ... },
  "timestamp": "ISO-8601"
}

Response Format:
{
  "id": "same-request-id",
  "success": true/false,
  "data": { ... },
  "error": "error-message-if-failed",
  "timestamp": "ISO-8601"
}
```

---

### 2. Модуль получения стен (GetWalls Module)

**Назначение**: Получение информации о стенах из Renga

**Компоненты**:
- `GetWallsCommand` - команда для получения стен
- `GetWallsHandler` - обработчик на стороне Renga
- `GetWallsComponent` - компонент Grasshopper

**Функции**:
- Запрос списка стен
- Парсинг геометрии стен
- Возврат данных в Grasshopper

**Команда**:
```json
{
  "id": "req-123",
  "command": "get_walls",
  "data": {},
  "timestamp": "2025-12-17T..."
}
```

**Ответ**:
```json
{
  "id": "req-123",
  "success": true,
  "data": {
    "walls": [
      {
        "id": 1,
        "name": "Wall 1",
        "position": { "x": 0, "y": 0, "z": 0 },
        "height": 3000,
        "thickness": 200,
        "baseline": { ... },
        "mesh": { ... }
      }
    ]
  },
  "timestamp": "2025-12-17T..."
}
```

---

### 3. Модуль создания колонн (CreateColumns Module)

**Назначение**: Создание и обновление колонн в Renga

**Компоненты**:
- `CreateColumnsCommand` - команда для создания колонн
- `CreateColumnsHandler` - обработчик на стороне Renga
- `CreateColumnsComponent` - компонент Grasshopper

**Функции**:
- Создание новых колонн
- Обновление существующих колонн
- Управление соответствием GUID

**Команда**:
```json
{
  "id": "req-456",
  "command": "update_points",
  "data": {
    "points": [
      {
        "x": 0.0,
        "y": 0.0,
        "z": 0.0,
        "height": 3000.0,
        "grasshopperGuid": "guid-1",
        "rengaColumnGuid": null
      }
    ]
  },
  "timestamp": "2025-12-17T..."
}
```

**Ответ**:
```json
{
  "id": "req-456",
  "success": true,
  "data": {
    "results": [
      {
        "success": true,
        "message": "Column created",
        "columnId": "123",
        "grasshopperGuid": "guid-1"
      }
    ]
  },
  "timestamp": "2025-12-17T..."
}
```

---

## Реализация TCP протокола

### Проблема текущей реализации

1. **Чтение данных**: Сервер пытается определить конец JSON по закрывающим скобкам, но это ненадежно
2. **Отправка ответа**: Ответ отправляется, но клиент может не успеть его прочитать
3. **Таймауты**: Неправильная обработка таймаутов

### Решение

#### 1. Протокол с длиной сообщения

**Формат сообщения**:
```
[4 байта: длина JSON в big-endian][JSON данные]
```

**Преимущества**:
- Точное определение конца сообщения
- Нет проблем с парсингом JSON
- Работает с любым размером данных

#### 2. Синхронная отправка/получение

**На стороне клиента**:
1. Отправить запрос
2. Сразу начать читать ответ
3. Прочитать длину (4 байта)
4. Прочитать JSON (N байт)
5. Закрыть соединение

**На стороне сервера**:
1. Принять соединение
2. Прочитать длину (4 байта)
3. Прочитать JSON (N байт)
4. Обработать команду
5. Отправить ответ (с длиной)
6. Закрыть соединение

---

## Структура файлов

### Grasshopper (Client Side)

```
GrasshopperRNG/
├── Connection/
│   ├── RengaConnectionClient.cs      # TCP клиент
│   ├── ConnectionProtocol.cs         # Протокол обмена
│   └── ConnectionMessage.cs          # Классы сообщений
├── Commands/
│   ├── GetWallsCommand.cs            # Команда получения стен
│   └── CreateColumnsCommand.cs       # Команда создания колонн
├── Components/
│   ├── RengaConnectComponent.cs      # Компонент связи
│   ├── RengaGetWallsComponent.cs     # Компонент получения стен
│   └── RengaCreateColumnsComponent.cs # Компонент создания колонн
└── Client/
    └── RengaGhClient.cs              # Старый клиент (заменить)
```

### Renga (Server Side)

```
RengaGH/
├── Connection/
│   ├── RengaConnectionServer.cs      # TCP сервер
│   ├── ConnectionProtocol.cs         # Протокол обмена
│   └── ConnectionMessage.cs          # Классы сообщений
├── Handlers/
│   ├── GetWallsHandler.cs            # Обработчик получения стен
│   └── CreateColumnsHandler.cs       # Обработчик создания колонн
├── Commands/
│   ├── ICommandHandler.cs            # Интерфейс обработчика
│   └── CommandRouter.cs              # Маршрутизатор команд
└── RengaPlugin.cs                    # Главный плагин
```

---

## План реализации

### Этап 1: Модуль связи (Connection)

1. ✅ Создать `ConnectionProtocol` с форматом длины сообщения
2. ✅ Переделать `RengaConnectionClient` с правильным чтением
3. ✅ Переделать `RengaConnectionServer` с правильной отправкой
4. ✅ Протестировать базовую связь

### Этап 2: Модуль получения стен (GetWalls)

1. ✅ Создать `GetWallsCommand` и `GetWallsHandler`
2. ✅ Интегрировать в `CommandRouter`
3. ✅ Обновить `RengaGetWallsComponent`
4. ✅ Протестировать получение стен

### Этап 3: Модуль создания колонн (CreateColumns)

1. ✅ Создать `CreateColumnsCommand` и `CreateColumnsHandler`
2. ✅ Интегрировать в `CommandRouter`
3. ✅ Обновить `RengaCreateColumnsComponent`
4. ✅ Протестировать создание колонн

---

## Детали реализации протокола

### ConnectionProtocol.cs (Client & Server)

```csharp
public static class ConnectionProtocol
{
    // Отправить сообщение с длиной
    public static async Task SendMessageAsync(NetworkStream stream, string json)
    {
        var data = Encoding.UTF8.GetBytes(json);
        var length = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(data.Length));
        
        await stream.WriteAsync(length, 0, 4);
        await stream.WriteAsync(data, 0, data.Length);
        await stream.FlushAsync();
    }
    
    // Получить сообщение с длиной
    public static async Task<string> ReceiveMessageAsync(NetworkStream stream, int timeoutMs = 10000)
    {
        // Читаем длину (4 байта)
        var lengthBytes = new byte[4];
        int totalRead = 0;
        while (totalRead < 4)
        {
            var read = await stream.ReadAsync(lengthBytes, totalRead, 4 - totalRead);
            if (read == 0) throw new IOException("Connection closed");
            totalRead += read;
        }
        
        int length = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(lengthBytes, 0));
        
        // Читаем JSON данные
        var buffer = new byte[length];
        totalRead = 0;
        while (totalRead < length)
        {
            var read = await stream.ReadAsync(buffer, totalRead, length - totalRead);
            if (read == 0) throw new IOException("Connection closed");
            totalRead += read;
        }
        
        return Encoding.UTF8.GetString(buffer, 0, length);
    }
}
```

---

## Тестирование

### Тест 1: Базовая связь
1. Запустить сервер в Renga
2. Подключиться из Grasshopper
3. Отправить тестовое сообщение
4. Получить ответ

### Тест 2: Получение стен
1. Создать стены в Renga
2. Запросить стены из Grasshopper
3. Проверить получение данных

### Тест 3: Создание колонн
1. Отправить точки из Grasshopper
2. Проверить создание колонн в Renga
3. Обновить точки
4. Проверить обновление колонн

---

## Миграция

1. Сохранить старые файлы как `.old`
2. Создать новые модули
3. Постепенно переносить функциональность
4. Тестировать каждый модуль отдельно
5. Удалить старые файлы после проверки

---

## Важные замечания

1. **Обратная совместимость**: Старые компоненты могут не работать, нужна миграция
2. **Логирование**: Добавить подробное логирование для отладки
3. **Обработка ошибок**: Четкая обработка всех ошибок
4. **Производительность**: Оптимизировать для больших объемов данных

