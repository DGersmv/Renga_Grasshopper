# План разработки: Получение свойств стен из Renga

## Цель
Расширить функциональность ноды `GetWalls` для получения дополнительных свойств стен и создать новую ноду для их отображения.

---

## Этап 1: Изучение API Renga для свойств стен ✅ ЗАВЕРШЕНО

### ⚠️ Важное примечание о параметрах
**Параметры** (`ParameterIds.*`) - это значения, которые можно использовать для **создания или изменения** стены. Они отличаются от свойств (`IWallParams.*`), которые являются read-only свойствами объекта.

Для получения значений параметров:
- Использовать `IModelObject.GetParameters().Get(ParameterIds.WallThickness)`
- Получить значение через `parameter.GetDoubleValue()` или `parameter.GetIntValue()`

### 1.1 Найденные интерфейсы и методы

#### ✅ **IWallParams** - основной интерфейс для параметров стены
- `Height` (property) - высота стены
- `Thickness` (property) - толщина стены
- `VerticalOffset` (property) - вертикальное смещение объекта
- `GetContour()` → `IWallContour` - получить контур стены

#### ✅ **IWallContour** - контур стены
- `GetBaseline()` → `ICurve2D` - линия привязки (baseline)
- `GetLeftCurve()` → `ICurve2D` - левая кривая контура
- `GetRightCurve()` → `ICurve2D` - правая кривая контура
- `GetCenterLine()` → `ICurve2D` - центральная линия контура

**Определение положения линии привязки**: Сравнить baseline с left/center/right кривыми

#### ✅ **IObjectWithLayeredMaterial** - интерфейс для объектов с многослойным материалом
- `HasLayeredMaterial()` → `bool` - есть ли многослойный материал
- `GetLayers()` → `ILayerCollection` - получить коллекцию слоев
- `LayeredMaterialId` (property) - ID многослойного материала

#### ✅ **IMaterialLayerCollection** - коллекция слоев материала
- `Count` (property) - количество слоев
- `Get(int index)` → `IMaterialLayer` - получить слой по индексу

#### ✅ **IMaterialLayer** - один слой материала
- `Thickness` (property) - толщина слоя
- `Material` (property) → `IMaterial` - материал слоя

#### ✅ **IMaterial** - материал
- `Name` (property) - имя материала
- `Id` (property) - ID материала

#### ✅ **ILevelObject** - интерфейс для объектов на уровне
- `GetLevel()` → `ILevel` - получить уровень (этаж)
- `LevelId` (property) - ID уровня
- `ElevationAboveLevel` (property) - высота объекта над уровнем
- `PlacementElevation` (property) - высота размещения над уровнем
- `VerticalOffset` (property) - вертикальное смещение от размещения

#### ✅ **ILevel** - уровень (этаж)
- `LevelName` (property) - имя уровня
- `Elevation` (property) - высота уровня относительно XOY плоскости

#### ✅ **IModelObject** - базовый интерфейс объекта
- `GetParameters()` → `IParameterContainer` - уже используется
- `ParameterIds.WallHeight` - высота стены (параметр)
- `ParameterIds.WallThickness` - толщина стены (параметр)
- `ParameterIds.WallPositionRelativeToBaseline` - положение линии привязки (параметр: 0=left, 1=center, 2=right)
- `ParameterIds.WallHorizontalOffset` - смещение от линии привязки (параметр)

### 1.2 Документирование найденных методов ✅
Все методы найдены и задокументированы выше.

---

## Этап 2: Расширение GetWallsHandler (Renga)

### 2.1 Создание метода ExtractWallProperties
Создать новый метод `ExtractWallProperties` в `GetWallsHandler.cs`:

```csharp
private object ExtractWallProperties(
    Renga.IModelObject wallObj, 
    Renga.ILevelObject levelObject,
    Renga.IWallParams wallParams,
    Renga.IApplication app)  // Добавить для доступа к LayeredMaterialManager
{
    var properties = new Dictionary<string, object>();
    
    // 1. Многослойная структура и материалы
    var objectWithLayeredMaterial = wallObj as Renga.IObjectWithLayeredMaterial;
    bool hasMultilayer = false;
    var layers = new List<object>();
    double totalThickness = 0;
    var materials = new List<string>();
    
    if (objectWithLayeredMaterial != null)
    {
        try
        {
                var layeredMaterialId = objectWithLayeredMaterial.LayeredMaterialId;
                if (layeredMaterialId > 0)
                {
                    var project = app.Project;
                    var layeredMaterialManager = project.LayeredMaterialManager;
                var layeredMaterial = layeredMaterialManager.GetLayeredMaterial(layeredMaterialId);
                
                if (layeredMaterial != null)
                {
                    hasMultilayer = true;
                    var layerCollection = layeredMaterial.Layers;
                    int layerCount = layerCollection.Count;
                    
                    for (int i = 0; i < layerCount; i++)
                    {
                        var layer = layerCollection.Get(i);
                        double layerThickness = layer.Thickness;
                        string materialName = layer.Material?.Name ?? "Unknown";
                        
                        totalThickness += layerThickness;
                        materials.Add(materialName);
                        
                        layers.Add(new Dictionary<string, object>
                        {
                            { "index", i },
                            { "thickness", layerThickness },
                            { "material", materialName }
                        });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            LogToFile($"Error getting layered material: {ex.Message}");
        }
    }
    
    // Если нет многослойной структуры, получить толщину из параметра
    if (!hasMultilayer)
    {
        try
        {
            var parameters = wallObj.GetParameters();
            if (parameters != null)
            {
                var thicknessParam = parameters.Get(Renga.ParameterIds.WallThickness);
                if (thicknessParam != null)
                {
                    totalThickness = thicknessParam.GetDoubleValue();
                }
                else
                {
                    // Fallback на IWallParams.Thickness
                    totalThickness = wallParams.Thickness;
                }
            }
        }
        catch
        {
            totalThickness = wallParams.Thickness;
        }
    }
    
    properties["hasMultilayerStructure"] = hasMultilayer;
    properties["layers"] = layers;
    properties["totalThickness"] = totalThickness;
    properties["materials"] = materials;
    
    // 2. Положение линии привязки и смещение (из параметров)
    string alignmentPosition = "Unknown";
    double alignmentOffset = 0.0;
    
    try
    {
        var parameters = wallObj.GetParameters();
        if (parameters != null)
        {
            // Положение линии привязки
            var alignmentParam = parameters.Get(Renga.ParameterIds.WallPositionRelativeToBaseline);
            if (alignmentParam != null)
            {
                int alignmentValue = alignmentParam.GetIntValue();
                switch (alignmentValue)
                {
                    case 0: alignmentPosition = "Left"; break;
                    case 1: alignmentPosition = "Center"; break;
                    case 2: alignmentPosition = "Right"; break;
                    case 3: alignmentPosition = "TopLeft"; break;
                    case 4: alignmentPosition = "TopCenter"; break;
                    case 5: alignmentPosition = "TopRight"; break;
                    case 6: alignmentPosition = "BottomLeft"; break;
                    case 7: alignmentPosition = "BottomCenter"; break;
                    case 8: alignmentPosition = "BottomRight"; break;
                    case 9: alignmentPosition = "CenterOfMass"; break;
                    case 10: alignmentPosition = "BasePoint"; break;
                    default: alignmentPosition = $"Unknown({alignmentValue})"; break;
                }
            }
            
            // Смещение от линии привязки
            var offsetParam = parameters.Get(Renga.ParameterIds.WallHorizontalOffset);
            if (offsetParam != null)
            {
                alignmentOffset = offsetParam.GetDoubleValue();
            }
        }
    }
    catch (Exception ex)
    {
        LogToFile($"Error getting alignment parameters: {ex.Message}");
    }
    
    properties["alignmentLinePosition"] = alignmentPosition;
    properties["alignmentOffset"] = alignmentOffset;
    
    // 3. Уровень (этаж)
    int levelId = levelObject.LevelId;
    string levelName = "Unknown";
    double levelElevation = 0.0;
    
    try
    {
        var level = levelObject.GetLevel();
        if (level != null)
        {
            levelName = level.LevelName ?? "Unknown";
            levelElevation = level.Elevation;
        }
    }
    catch { }
    
    properties["levelId"] = levelId;
    properties["levelName"] = levelName;
    
    // 4. Смещение от уровня этажа
    double levelOffset = levelObject.ElevationAboveLevel;
    properties["levelOffset"] = levelOffset;
    
    // 5. Высота стены (из параметра)
    double height = 0.0;
    try
    {
        var parameters = wallObj.GetParameters();
        if (parameters != null)
        {
            var heightParam = parameters.Get(Renga.ParameterIds.WallHeight);
            if (heightParam != null)
            {
                height = heightParam.GetDoubleValue();
            }
            else
            {
                // Fallback на IWallParams.Height
                height = wallParams.Height;
            }
        }
        else
        {
            height = wallParams.Height;
        }
    }
    catch
    {
        height = wallParams.Height;
    }
    properties["height"] = height;
    
    return properties;
}
```

### 2.2 Реализация получения свойств

#### 2.2.1 Многослойная структура ✅
- [x] Проверить наличие многослойной структуры через `IObjectWithLayeredMaterial`
- [x] Если есть:
  - Получить `LayeredMaterialId` через `objectWithLayeredMaterial.LayeredMaterialId`
  - Через `LayeredMaterialManager.GetLayeredMaterial(LayeredMaterialId)` получить `ILayeredMaterial`
  - Получить количество слоев через `layeredMaterial.Layers.Count`
  - Для каждого слоя получить:
    - Толщину слоя: `layer.Thickness`
    - Материал слоя: `layer.Material.Name`
  - Общая толщина = сумма толщин слоев
- [x] Если нет:
  - Использовать параметр `ParameterIds.WallThickness` (значение параметра для создания/изменения)

#### 2.2.2 Материалы ✅
- [x] Получить материалы через `IObjectWithLayeredMaterial`
- [x] Если многослойная структура:
  - Материалы каждого слоя: `layer.Material.Name`
- [x] Если нет:
  - Возможно через параметры (нужно проверить на практике)

#### 2.2.3 Положение линии привязки ✅
- [x] Использовать параметр `ParameterIds.WallPositionRelativeToBaseline`
- [x] Значения параметра:
  - 0 : Relative to left → "Left"
  - 1 : Relative to center → "Center"
  - 2 : Relative to right → "Right"
  - 3-10 : Другие варианты (top left, bottom right, etc.)
- [x] Получить через `parameters.Get(ParameterIds.WallPositionRelativeToBaseline).GetIntValue()`

#### 2.2.4 Смещение от линии привязки ✅
- [x] Использовать параметр `ParameterIds.WallHorizontalOffset`
- [x] Получить через `parameters.Get(ParameterIds.WallHorizontalOffset).GetDoubleValue()`
- [x] Это значение параметра, которое можно использовать для создания/изменения стены

#### 2.2.5 Уровень (этаж) ✅
- [x] Использовать `ILevelObject.GetLevel()` → `ILevel`
- [x] Получить:
  - ID уровня: `ILevelObject.LevelId`
  - Имя уровня: `ILevel.LevelName`
  - Высоту уровня: `ILevel.Elevation`

#### 2.2.6 Смещение от уровня этажа ✅
- [x] Использовать `ILevelObject.ElevationAboveLevel` (высота объекта над уровнем)
- [x] Или `ILevelObject.PlacementElevation` (высота размещения над уровнем)
- [x] Или вычислить: `placement.Origin.Z - level.Elevation`

#### 2.2.7 Высота стены ✅
- [x] Использовать параметр `ParameterIds.WallHeight` (значение параметра для создания/изменения)
- [x] Или `IWallParams.Height` (property) как fallback

### 2.3 Интеграция в GetWallsHandler.Handle
- [ ] Вызвать `ExtractWallProperties` для каждой стены
- [ ] Передать `m_app` для доступа к `LayeredMaterialManager`
- [ ] Добавить `properties` в объект `wallData`:

```csharp
var wallData = new
{
    id = obj.Id,
    name = obj.Name ?? $"Wall {obj.Id}",
    position = ...,
    height = height,
    thickness = thickness,
    baseline = baselineData,
    mesh = meshData,
    properties = ExtractWallProperties(obj, wall, wallParams, m_app)  // НОВОЕ: добавить m_app
};
```

---

## Этап 3: Расширение RengaGetWallsComponent (Grasshopper)

### 3.1 Добавление нового выхода "Properties"
- [ ] В методе `RegisterOutputParams` добавить:
  ```csharp
  pManager.AddGenericParameter("Properties", "P", "Wall properties", GH_ParamAccess.list);
  ```

### 3.2 Создание класса WallProperties
- [ ] Создать класс `WallProperties` для хранения свойств:
  ```csharp
  public class WallProperties
  {
      public int WallId { get; set; }
      public string WallName { get; set; }
      
      // Многослойная структура
      public bool HasMultilayerStructure { get; set; }
      public List<LayerInfo> Layers { get; set; }  // Если многослойная
      public double Thickness { get; set; }  // Общая толщина
      
      // Материалы
      public List<string> Materials { get; set; }
      
      // Линия привязки
      public string AlignmentLinePosition { get; set; }  // "Left", "Center", "Right"
      public double AlignmentOffset { get; set; }
      
      // Уровень
      public int LevelId { get; set; }
      public string LevelName { get; set; }
      public double LevelOffset { get; set; }
      
      // Высота
      public double Height { get; set; }
  }
  
  public class LayerInfo
  {
      public int Index { get; set; }
      public double Thickness { get; set; }
      public string Material { get; set; }
  }
  ```

### 3.3 Парсинг свойств в SolveInstance
- [ ] В методе `SolveInstance` добавить парсинг `properties`:
  ```csharp
  var propertiesList = new List<WallProperties>();
  
  foreach (var wallToken in walls)
  {
      var wall = wallToken as JObject;
      var properties = ParseWallProperties(wall);
      if (properties != null)
          propertiesList.Add(properties);
  }
  
  DA.SetDataList(4, propertiesList);  // Новый выход
  ```

### 3.4 Создание метода ParseWallProperties
- [ ] Реализовать парсинг JSON свойств в объект `WallProperties`

---

## Этап 4: Создание новой ноды WallPropertiesComponent

### 4.1 Создание компонента
- [ ] Создать файл `RengaWallPropertiesComponent.cs`
- [ ] Наследовать от `GH_Component`
- [ ] Установить уникальный GUID
- [ ] Настроить категорию: "Renga", "Walls"

### 4.2 Определение входов и выходов

#### Входы:
- [ ] `Properties` (Generic) - список `WallProperties` от `GetWalls`

#### Выходы:
- [ ] `Wall IDs` (Integer) - ID стен
- [ ] `Wall Names` (Text) - имена стен
- [ ] `Has Multilayer` (Boolean) - есть ли многослойная структура
- [ ] `Layer Count` (Integer) - количество слоев (0 если нет многослойной)
- [ ] `Layer Thicknesses` (Number) - толщины слоев (список списков)
- [ ] `Layer Materials` (Text) - материалы слоев (список списков)
- [ ] `Total Thickness` (Number) - общая толщина стены
- [ ] `Materials` (Text) - все материалы стены (список)
- [ ] `Alignment Position` (Text) - положение линии привязки
- [ ] `Alignment Offset` (Number) - смещение от линии привязки
- [ ] `Level ID` (Integer) - ID уровня
- [ ] `Level Name` (Text) - имя уровня
- [ ] `Level Offset` (Number) - смещение от уровня
- [ ] `Height` (Number) - высота стены

### 4.3 Реализация SolveInstance
- [ ] Получить список `WallProperties` из входа
- [ ] Распределить данные по соответствующим выходам
- [ ] Обработать многослойные структуры (списки списков)
- [ ] Обработать случаи, когда свойство отсутствует

### 4.4 Создание атрибутов компонента
- [ ] Создать `RengaWallPropertiesComponentAttributes.cs`
- [ ] Настроить отображение компонента

---

## Этап 5: Тестирование

### 5.1 Тестирование получения свойств
- [ ] Создать тестовую модель в Renga с разными типами стен:
  - Простая стена (без многослойной структуры)
  - Многослойная стена
  - Стены с разными положениями линии привязки
  - Стены на разных уровнях
  - Стены с разными смещениями

### 5.2 Проверка данных
- [ ] Проверить, что все свойства корректно извлекаются
- [ ] Проверить логи в `RengaGH_GetWalls.log`
- [ ] Проверить, что данные корректно передаются в Grasshopper

### 5.3 Проверка новой ноды
- [ ] Подключить `WallPropertiesComponent` к выходу `Properties` из `GetWalls`
- [ ] Проверить, что все выходы отображают корректные данные
- [ ] Проверить обработку многослойных структур
- [ ] Проверить обработку отсутствующих свойств

---

## Этап 6: Обработка ошибок и edge cases

### 6.1 Обработка отсутствующих данных
- [ ] Если многослойная структура отсутствует - использовать толщину стены
- [ ] Если материал отсутствует - возвращать "Unknown" или пустую строку
- [ ] Если уровень отсутствует - возвращать null или -1
- [ ] Если alignment неизвестен - возвращать "Unknown"

### 6.2 Логирование
- [ ] Добавить логирование в `ExtractWallProperties`
- [ ] Логировать случаи, когда свойство недоступно
- [ ] Логировать ошибки при извлечении свойств

---

## Известные вопросы для исследования ✅ РЕШЕНО

### ✅ Вопрос 1: Многослойная структура - РЕШЕНО
- **Ответ**: Использовать `IObjectWithLayeredMaterial.HasLayeredMaterial()` и `GetLayers()`
- **Материалы слоев**: `IMaterialLayer.Material.Name`

### ✅ Вопрос 2: Положение линии привязки - РЕШЕНО
- **Ответ**: Сравнить `IWallContour.GetBaseline()` с `GetLeftCurve()`, `GetCenterLine()`, `GetRightCurve()`
- **Значения**: "Left", "Center", "Right" (определяется через сравнение кривых)

### ✅ Вопрос 3: Смещения - РЕШЕНО
- **Смещение от линии привязки**: Вычислить расстояние между baseline и left/right/center
- **Смещение от уровня**: Использовать `ILevelObject.ElevationAboveLevel` или `PlacementElevation`

### ✅ Вопрос 4: Материалы - РЕШЕНО
- **Ответ**: Из многослойной структуры через `IObjectWithLayeredMaterial.GetLayers()`
- **Для каждого слоя**: `IMaterialLayer.Material.Name`

---

## Следующие шаги

1. **Исследовать API Renga** - найти все необходимые методы
2. **Реализовать ExtractWallProperties** в GetWallsHandler
3. **Добавить выход Properties** в RengaGetWallsComponent
4. **Создать WallPropertiesComponent** для отображения свойств
5. **Протестировать** на реальной модели

---

## Структура данных Properties (JSON)

```json
{
  "wallId": 100014,
  "wallName": "Стена: 200,00 мм",
  "hasMultilayerStructure": true,
  "layers": [
    {
      "index": 0,
      "thickness": 50.0,
      "material": "Кирпич"
    },
    {
      "index": 1,
      "thickness": 100.0,
      "material": "Утеплитель"
    },
    {
      "index": 2,
      "thickness": 50.0,
      "material": "Кирпич"
    }
  ],
  "totalThickness": 200.0,
  "materials": ["Кирпич", "Утеплитель", "Кирпич"],
  "alignmentLinePosition": "Center",
  "alignmentOffset": 0.0,
  "levelId": 1,
  "levelName": "Этаж 1",
  "levelOffset": 0.0,
  "height": 3000.0
}
```

Или для простой стены (без многослойной структуры):

```json
{
  "wallId": 100015,
  "wallName": "Стена: 150,00 мм",
  "hasMultilayerStructure": false,
  "layers": null,
  "totalThickness": 150.0,
  "materials": ["Бетон"],
  "alignmentLinePosition": "Left",
  "alignmentOffset": 50.0,
  "levelId": 1,
  "levelName": "Этаж 1",
  "levelOffset": 100.0,
  "height": 3000.0
}
```

