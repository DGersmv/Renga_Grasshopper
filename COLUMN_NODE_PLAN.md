# План для ноды создания колонн в Renga (RengaCreateColumnsComponent)

## Обзор

Нода **"Renga Create Columns"** - это компонент Grasshopper, который принимает точки из Grasshopper и отправляет их в Renga для создания/обновления колонн через TCP соединение.

## Текущее состояние

### ✅ Уже реализовано:
- Базовая структура компонента `RengaCreateColumnsComponent`
- Интеграция с главной нодой `RengaConnectComponent`
- Отправка данных через TCP клиент
- Обработка ответов от сервера Renga
- Кастомные атрибуты компонента

### ⚠️ Требует доработки:
- Метод `GetPointGuid()` - не реализован (используется, но отсутствует)
- Кнопка "Update" в кастомных атрибутах - не реализована
- Отслеживание изменений координат точек
- Валидация входных данных
- Иконка компонента
- Обработка единиц измерения (если нужно)
- **Задание высоты для каждой колонны** - отсутствует (критично!)

---

## Детальный план реализации

### 1. Реализация метода GetPointGuid()

**Задача**: Создать уникальный идентификатор для каждой точки, чтобы отслеживать соответствие между точками Grasshopper и колоннами Renga.

**Варианты реализации**:

#### Вариант 1: Использование Rhino GUID (предпочтительно)
```csharp
private string GetPointGuid(Point3d point)
{
    // Если точка имеет Rhino GUID (из геометрии Rhino)
    if (point is GH_Point ghPoint && ghPoint.Value != null)
    {
        var rhinoPoint = ghPoint.Value;
        // Попытка получить GUID из геометрии Rhino
        // Если доступен - использовать его
    }
    
    // Если GUID недоступен - генерировать стабильный GUID на основе координат
    // ВАЖНО: Это должно быть стабильно при одинаковых координатах
    return GenerateStableGuid(point);
}
```

#### Вариант 2: Генерация стабильного GUID на основе координат
```csharp
private string GetPointGuid(Point3d point)
{
    // Использовать координаты для генерации стабильного GUID
    // Проблема: при изменении координат GUID изменится
    // Решение: хранить соответствие в статическом словаре с координатами
    return $"GH_Point_{point.X}_{point.Y}_{point.Z}";
}
```

#### Вариант 3: Использование индекса точки в списке (не рекомендуется)
- Проблема: при изменении порядка точек соответствие нарушится
- Не использовать для продакшена

**Рекомендация**: Комбинированный подход:
1. Попытаться получить Rhino GUID точки
2. Если недоступен - использовать статический словарь для хранения соответствия индекса/координат и GUID
3. Генерировать GUID один раз при первом появлении точки

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

### 2. Реализация кнопки "Update" в кастомных атрибутах

**Задача**: Добавить кнопку "Update" в компонент для принудительного обновления всех колонн.

**Шаги**:
1. Реализовать метод `Render()` в `RengaCreateColumnsComponentAttributes`
2. Добавить кнопку "Update" в визуальное представление компонента
3. Обработать клик по кнопке через `GH_Attributes.MouseDown()`
4. Вызвать метод `OnUpdateButtonClick()` компонента

**Пример структуры**:
```csharp
public class RengaCreateColumnsComponentAttributes : GH_ComponentAttributes
{
    private RectangleF updateButtonRect;
    
    protected override void Layout()
    {
        base.Layout();
        // Вычислить позицию кнопки
        updateButtonRect = new RectangleF(...);
    }
    
    protected override void Render(GH_Canvas canvas, Graphics graphics, GH_CanvasChannel channel)
    {
        if (channel == GH_CanvasChannel.Objects)
        {
            // Нарисовать кнопку "Update"
            // ...
        }
        base.Render(canvas, graphics, channel);
    }
    
    public override GH_ObjectResponse RespondToMouseDown(GH_Canvas sender, GH_CanvasMouseEvent e)
    {
        if (updateButtonRect.Contains(e.CanvasLocation))
        {
            (Owner as RengaCreateColumnsComponent)?.OnUpdateButtonClick();
            return GH_ObjectResponse.Handled;
        }
        return base.RespondToMouseDown(sender, e);
    }
}
```

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponentAttributes.cs`

---

### 3. Отслеживание изменений координат точек

**Задача**: Определять, изменились ли координаты точки, чтобы отправлять обновления только при реальных изменениях.

**Реализация**:
1. Хранить предыдущие координаты точек в статическом словаре
2. При каждом решении сравнивать текущие координаты с предыдущими
3. Отправлять обновления только для точек с измененными координатами

**Структура данных**:
```csharp
private static Dictionary<string, Point3d> pointGuidToLastCoordinates = new Dictionary<string, Point3d>();

private bool HasCoordinatesChanged(string pointGuid, Point3d currentPoint)
{
    if (!pointGuidToLastCoordinates.ContainsKey(pointGuid))
    {
        // Новая точка
        pointGuidToLastCoordinates[pointGuid] = currentPoint;
        return true;
    }
    
    var lastPoint = pointGuidToLastCoordinates[pointGuid];
    var tolerance = 0.001; // Точность сравнения (в единицах Grasshopper)
    
    if (Math.Abs(currentPoint.X - lastPoint.X) > tolerance ||
        Math.Abs(currentPoint.Y - lastPoint.Y) > tolerance ||
        Math.Abs(currentPoint.Z - lastPoint.Z) > tolerance)
    {
        // Координаты изменились
        pointGuidToLastCoordinates[pointGuid] = currentPoint;
        return true;
    }
    
    return false; // Координаты не изменились
}
```

**Оптимизация**: Отправлять обновления только для точек с измененными координатами или новых точек.

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

### 4. Валидация входных данных

**Задача**: Проверять корректность входных данных перед отправкой в Renga.

**Проверки**:
1. **Точки**: 
   - Список не пустой
   - Все точки валидны (не null, не NaN, не Infinity)
   - Координаты в разумных пределах

2. **RengaConnect компонент**:
   - Компонент подключен
   - Клиент валиден и подключен к серверу

3. **Координаты**:
   - Проверка на NaN и Infinity
   - Проверка на разумные значения (например, в пределах ±1000000)

**Реализация**:
```csharp
private bool ValidateInputs(List<Point3d> points, RengaGhClient client, out string errorMessage)
{
    errorMessage = "";
    
    if (points == null || points.Count == 0)
    {
        errorMessage = "No points provided";
        return false;
    }
    
    foreach (var point in points)
    {
        if (point == null || 
            double.IsNaN(point.X) || double.IsInfinity(point.X) ||
            double.IsNaN(point.Y) || double.IsInfinity(point.Y) ||
            double.IsNaN(point.Z) || double.IsInfinity(point.Z))
        {
            errorMessage = $"Invalid point coordinates: ({point?.X}, {point?.Y}, {point?.Z})";
            return false;
        }
    }
    
    if (client == null || !client.IsConnected)
    {
        errorMessage = "Renga Connect is not connected";
        return false;
    }
    
    return true;
}
```

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

### 5. Добавление иконки компонента

**Задача**: Создать и добавить иконку для компонента.

**Шаги**:
1. Создать иконку (24x24 или 32x32 пикселей, PNG формат)
2. Добавить файл иконки в проект (например, `Resources/ColumnIcon.png`)
3. Установить свойство `Build Action = Embedded Resource`
4. Загрузить иконку в свойстве `Icon` компонента

**Реализация**:
```csharp
protected override Bitmap Icon
{
    get
    {
        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
        var resourceName = "GrasshopperRNG.Resources.ColumnIcon.png";
        using (var stream = assembly.GetManifestResourceStream(resourceName))
        {
            if (stream != null)
            {
                return new Bitmap(stream);
            }
        }
        // Fallback: вернуть стандартную иконку или null
        return null;
    }
}
```

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

### 6. Задание высоты для каждой колонны ⚠️ КРИТИЧНО

**Задача**: Добавить возможность задавать высоту для каждой колонны из Grasshopper.

**Текущая проблема**: 
- В компоненте нет параметра для высоты колонн
- В JSON формате нет поля `height`
- В Renga плагине не устанавливается высота при создании колонны

**Решение**:

#### 6.1 Добавить параметр высоты в компонент Grasshopper

**Варианты**:
1. **Одно значение для всех колонн** (проще):
   - Параметр `Height` (Number) - одно значение высоты для всех колонн
   - По умолчанию: 3000 мм (3 метра) или значение из настроек Renga

2. **Список высот для каждой колонны** (гибче):
   - Параметр `Heights` (Number, List) - список высот, соответствующий списку точек
   - Если список короче списка точек - использовать последнее значение для остальных
   - Если список пустой - использовать значение по умолчанию

**Рекомендация**: Начать с варианта 1 (одно значение), затем добавить вариант 2 (список).

**Реализация в RegisterInputParams**:
```csharp
// Вариант 1: Одно значение для всех
pManager.AddNumberParameter("Height", "H", "Column height in millimeters (default: 3000)", GH_ParamAccess.item, 3000.0);

// Вариант 2: Список высот (более гибко)
pManager.AddNumberParameter("Heights", "H", "Column heights in millimeters (one per point, or single value for all)", GH_ParamAccess.list);
```

**Реализация в SolveInstance**:
```csharp
// Вариант 1: Одно значение
double height = 3000.0; // значение по умолчанию
if (DA.GetData(2, ref height))
{
    if (height <= 0)
    {
        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Height must be positive. Using default 3000mm");
        height = 3000.0;
    }
}

// Вариант 2: Список высот
List<double> heights = new List<double>();
DA.GetDataList(2, heights);

// Нормализовать список: если меньше точек - дополнить последним значением
while (heights.Count < points.Count)
{
    if (heights.Count > 0)
        heights.Add(heights.Last());
    else
        heights.Add(3000.0); // значение по умолчанию
}
```

#### 6.2 Добавить поле height в JSON формат

**Изменения в PrepareCommand**:
```csharp
private object PrepareCommand(List<Point3d> points, List<double> heights)
{
    var pointData = new List<object>();

    for (int i = 0; i < points.Count; i++)
    {
        var point = points[i];
        var height = i < heights.Count ? heights[i] : heights.LastOrDefault(3000.0);
        
        var pointGuid = GetPointGuid(point);
        var rengaColumnGuid = pointGuidToColumnGuidMap.ContainsKey(pointGuid) 
            ? pointGuidToColumnGuidMap[pointGuid] 
            : null;

        pointData.Add(new
        {
            x = point.X,
            y = point.Y,
            z = point.Z,
            height = height,  // ← ДОБАВИТЬ
            grasshopperGuid = pointGuid,
            rengaColumnGuid = rengaColumnGuid
        });
    }

    return new
    {
        command = "update_points",
        points = pointData,
        timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    };
}
```

**Новый формат JSON**:
```json
{
  "command": "update_points",
  "points": [
    {
      "x": 0.0,
      "y": 0.0,
      "z": 0.0,
      "height": 3000.0,
      "grasshopperGuid": "guid-1",
      "rengaColumnGuid": null
    },
    {
      "x": 1000.0,
      "y": 2000.0,
      "z": 0.0,
      "height": 4000.0,
      "grasshopperGuid": "guid-2",
      "rengaColumnGuid": null
    }
  ],
  "timestamp": "2025-01-XX..."
}
```

#### 6.3 Установка высоты в Renga плагине

**Изменения в ProcessPoint**:
```csharp
private PointResult ProcessPoint(JObject? pointObj)
{
    // ... существующий код ...
    
    var x = pointObj["x"]?.Value<double>() ?? 0;
    var y = pointObj["y"]?.Value<double>() ?? 0;
    var z = pointObj["z"]?.Value<double>() ?? 0;
    var height = pointObj["height"]?.Value<double>() ?? 3000.0; // ← ДОБАВИТЬ
    var grasshopperGuid = pointObj["grasshopperGuid"]?.ToString();
    var rengaColumnGuid = pointObj["rengaColumnGuid"]?.ToString();
    
    // ... остальной код ...
    
    if (columnExists)
    {
        return UpdateColumn(columnId, x, y, z, height, grasshopperGuid); // ← добавить height
    }
    else
    {
        return CreateColumn(x, y, z, height, grasshopperGuid); // ← добавить height
    }
}
```

**Изменения в CreateColumn**:
```csharp
private PointResult CreateColumn(double x, double y, double z, double height, string grasshopperGuid)
{
    // ... существующий код создания колонны ...
    
    // После установки placement, установить высоту
    // ВАЖНО: Нужно найти правильный API метод для установки высоты колонны в Renga
    // Возможные варианты:
    
    // Вариант 1: Через IPropertyContainer (если доступен)
    if (column is IPropertyContainer propContainer)
    {
        // Попытка установить высоту через свойства
        // Требует знания имени свойства высоты в Renga
        // Например: propContainer.SetProperty("Height", height);
    }
    
    // Вариант 2: Через изменение размещения (если высота = расстояние до следующего уровня)
    // Установить конечную точку колонны на высоте height от основания
    var placement = column.GetPlacement();
    if (placement != null)
    {
        // Высота колонны может быть задана через изменение оси Z
        // Или через установку верхней точки колонны
        // Требует проверки API Renga
    }
    
    // Вариант 3: Через специальный метод SetHeight (если существует)
    // if (column is IColumn columnInterface)
    // {
    //     columnInterface.SetHeight(height);
    // }
    
    // ВАЖНО: Нужно проверить документацию Renga API или примеры для правильного метода
    // Временное решение: логировать высоту, но не устанавливать (до уточнения API)
    System.Diagnostics.Debug.WriteLine($"Column height requested: {height}mm for GUID: {grasshopperGuid}");
    
    // ... остальной код ...
}
```

**Изменения в UpdateColumn**:
```csharp
private PointResult UpdateColumn(int columnId, double x, double y, double z, double height, string grasshopperGuid)
{
    // ... существующий код обновления позиции ...
    
    // Также обновить высоту колонны
    // (аналогично CreateColumn, найти правильный API метод)
    
    // ... остальной код ...
}
```

**Примечание**: 
- Требуется уточнить API Renga для установки высоты колонны
- Возможные интерфейсы: `IPropertyContainer`, `IColumn`, или специальные методы `ILevelObject`
- Если API недоступен напрямую, может потребоваться установка через свойства объекта

**Файлы для изменения**:
1. `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs` - добавить параметр высоты
2. `RengaGH/RengaPlugin.cs` - добавить обработку высоты в JSON и установку в Renga

---

### 7. Обработка единиц измерения (опционально)

**Задача**: Конвертировать единицы измерения между Grasshopper и Renga (если необходимо).

**Проблема**: 
- Grasshopper может работать в метрах
- Renga работает в миллиметрах

**Решение**:
1. Добавить параметр компонента для выбора единиц измерения входных точек
2. Конвертировать координаты перед отправкой в Renga

**Реализация**:
```csharp
public enum InputUnits
{
    Millimeters,  // По умолчанию (как в Renga)
    Meters,
    Centimeters
}

// В RegisterInputParams:
pManager.AddIntegerParameter("Units", "U", "Input units: 0=mm, 1=m, 2=cm", GH_ParamAccess.item, 0);

// В PrepareCommand:
private double ConvertToMillimeters(double value, InputUnits units)
{
    switch (units)
    {
        case InputUnits.Meters:
            return value * 1000.0;
        case InputUnits.Centimeters:
            return value * 10.0;
        case InputUnits.Millimeters:
        default:
            return value;
    }
}
```

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

### 8. Улучшение обработки ошибок

**Задача**: Улучшить обработку и отображение ошибок пользователю.

**Улучшения**:
1. Более детальные сообщения об ошибках
2. Логирование ошибок (опционально)
3. Визуальная индикация ошибок (цвет компонента)

**Реализация**:
```csharp
// В SolveInstance:
try
{
    // ... основная логика
}
catch (JsonException ex)
{
    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
        $"JSON serialization error: {ex.Message}");
}
catch (SocketException ex)
{
    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
        $"Network error: {ex.Message}");
}
catch (Exception ex)
{
    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
        $"Unexpected error: {ex.Message}");
}
```

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

### 9. Оптимизация производительности

**Задача**: Оптимизировать работу компонента для больших количеств точек.

**Оптимизации**:
1. **Батчинг**: Отправлять точки батчами (например, по 100 точек за раз)
2. **Кэширование**: Кэшировать результаты предыдущих запросов
3. **Ленивая отправка**: Отправлять обновления только при реальных изменениях
4. **Асинхронность**: Использовать асинхронные методы для отправки данных (если возможно)

**Реализация батчинга**:
```csharp
private const int BatchSize = 100;

private void SendPointsInBatches(List<Point3d> points, RengaGhClient client)
{
    for (int i = 0; i < points.Count; i += BatchSize)
    {
        var batch = points.Skip(i).Take(BatchSize).ToList();
        var command = PrepareCommand(batch);
        var json = JsonConvert.SerializeObject(command);
        var responseJson = client.Send(json);
        // Обработать ответ для батча
    }
}
```

**Файл для изменения**: `GrasshopperRNG/Components/RengaCreateColumnsComponent.cs`

---

## Порядок реализации

### Этап 1: Критичные функции (обязательно)
1. ✅ Реализовать метод `GetPointGuid()` - **КРИТИЧНО** (компонент не работает без него)
2. ✅ **Добавить задание высоты для каждой колонны** - **КРИТИЧНО** (требование пользователя)
3. ✅ Добавить валидацию входных данных
4. ✅ Улучшить обработку ошибок
5. ✅ **Сгенерировать уникальный GUID для компонента** - **КРИТИЧНО** (каждый компонент должен иметь уникальный GUID)

### Этап 2: Улучшение UX (желательно)
5. ✅ Реализовать кнопку "Update" в кастомных атрибутах
6. ✅ Добавить иконку компонента
7. ✅ Реализовать отслеживание изменений координат

### Этап 3: Дополнительные функции (опционально)
8. ⚪ Добавить обработку единиц измерения
9. ⚪ Оптимизировать производительность для больших количеств точек

---

## Тестирование

### Тестовые сценарии:

1. **Базовый тест**:
   - Подключить компонент к RengaConnect
   - Добавить несколько точек
   - Нажать "Update"
   - Проверить создание колонн в Renga

2. **Тест обновления**:
   - Изменить координаты существующих точек
   - Нажать "Update"
   - Проверить обновление позиций колонн в Renga

3. **Тест добавления новых точек**:
   - Добавить новые точки к существующим
   - Нажать "Update"
   - Проверить создание колонн только для новых точек

4. **Тест удаления точек**:
   - Удалить некоторые точки из списка
   - Нажать "Update"
   - Проверить, что существующие колонны не удалены (требование из PLAN.md)

5. **Тест ошибок**:
   - Отключить RengaConnect
   - Попытаться отправить точки
   - Проверить сообщение об ошибке

6. **Тест валидации**:
   - Добавить точки с NaN координатами
   - Проверить сообщение об ошибке

---

## Файлы для изменения

1. **GrasshopperRNG/Components/RengaCreateColumnsComponent.cs**
   - Добавить метод `GetPointGuid()`
   - **Добавить параметр высоты колонн (Height/Heights)**
   - **Добавить поле height в PrepareCommand()**
   - **Сгенерировать и добавить уникальный ComponentGuid** (каждый компонент Grasshopper должен иметь уникальный GUID)
     - Использовать `[guid]::NewGuid()` в PowerShell или онлайн генератор GUID
     - Убедиться, что GUID уникален и не совпадает с другими компонентами
     - Добавить `public override Guid ComponentGuid => new Guid("...");` в класс компонента
   - Добавить метод `HasCoordinatesChanged()`
   - Добавить метод `ValidateInputs()`
   - Улучшить обработку ошибок
   - Добавить оптимизации (опционально)

2. **RengaGH/RengaPlugin.cs**
   - **Добавить обработку поля height в ProcessPoint()**
   - **Добавить параметр height в CreateColumn() и UpdateColumn()**
   - **Реализовать установку высоты колонны через Renga API**

3. **GrasshopperRNG/Components/RengaCreateColumnsComponentAttributes.cs**
   - Реализовать метод `Render()` для кнопки "Update"
   - Реализовать метод `RespondToMouseDown()` для обработки клика
   - Реализовать метод `Layout()` для позиционирования кнопки

4. **GrasshopperRNG/Resources/** (создать папку, если нужно)
   - Добавить файл иконки `ColumnIcon.png`

---

## Заметки и рекомендации

1. **GUID точек**: 
   - Важно использовать стабильный GUID, который не меняется при изменении координат
   - Рекомендуется использовать комбинированный подход (Rhino GUID + статический словарь)

2. **Производительность**:
   - Для больших количеств точек (>1000) рассмотреть батчинг
   - Кэшировать результаты предыдущих запросов

3. **Обработка ошибок**:
   - Всегда проверять подключение перед отправкой данных
   - Предоставлять понятные сообщения об ошибках пользователю

4. **Тестирование**:
   - Тестировать с различными количествами точек (1, 10, 100, 1000)
   - Тестировать сценарии добавления/удаления/изменения точек
   - Тестировать обработку ошибок

---

## Следующие шаги

1. **Начать с реализации метода `GetPointGuid()`** - это критично для работы компонента
2. **Добавить задание высоты для каждой колонны** - критичное требование:
   - Добавить параметр Height в компонент Grasshopper
   - Добавить поле height в JSON формат
   - Найти и использовать правильный API метод для установки высоты в Renga
3. **Сгенерировать уникальный GUID для компонента** - критично:
   - Использовать `[guid]::NewGuid()` в PowerShell или онлайн генератор GUID
   - Убедиться, что GUID уникален и не совпадает с другими компонентами
   - Добавить `public override Guid ComponentGuid => new Guid("...");` в класс компонента
4. Добавить валидацию входных данных
5. Реализовать кнопку "Update" в кастомных атрибутах
6. Протестировать базовую функциональность
7. Добавить дополнительные улучшения по мере необходимости

---

## Важные замечания по высоте колонн

### Необходимо уточнить:
1. **API метод для установки высоты** в Renga:
   - Проверить документацию Renga SDK
   - Проверить примеры использования API для колонн
   - Возможные варианты: `IPropertyContainer`, `IColumn`, или специальные методы

2. **Единицы измерения высоты**:
   - Renga работает в миллиметрах
   - Убедиться, что высота передается в правильных единицах
   - Возможно, нужна конвертация из метров (если Grasshopper в метрах)

3. **Валидация высоты**:
   - Минимальная высота (например, > 0)
   - Максимальная высота (если есть ограничения)
   - Проверка на разумные значения

