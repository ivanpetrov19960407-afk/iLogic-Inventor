# iLogic-Inventor

Migration workspace for moving legacy VBA Inventor automation into a modular C# library.

## Current .NET core library (Task 1)

`src/StoneAutomation` (assembly/root namespace: `StoneDrawingAuto.Core`)

- `Core/StoneMathService.cs` – unit conversion, Point2d helpers, and A3 landscape validation.
- `Models/ProjectItem.cs` – normalized order model for later Excel parsing.
- `Abstractions/` – placeholders for module contracts:
  - `IOrderProcessor`
  - `IDrawingAutomator`
  - `ISheetStylingService`
