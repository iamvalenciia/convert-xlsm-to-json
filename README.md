## XLSM TO JSON

high-performance JSON viewer, search engine, and workflow debugger designed for complex Excel-derived (`.xlsm`) configuration files. It transforms raw spreadsheet data into structured JSON, allowing developers and analysts to inspect, search, copy, and analyze workflows with significantly more clarity and speed than traditional Excel tools.

---

## Download for Windows

- [XLSM TO JSON Inspect v1.0.0 (Windows x64)](https://github.com/iamvalenciia/hammer-inspect/releases/download/v1.0.0/hammer-inspect.exe)
- [GitHub Releases Page](https://github.com/iamvalenciia/hammer-inspect/releases)

---

## Core Features

- Instant keyword searches across 30,000+ rows
- One-click `Copy Search` export for AI workflow analysis
- Structured and readable JSON visualization
- GPU-accelerated rendering for smooth scrolling and UI performance
- Native Windows UI with modern WinUI dialogs and controls
- Self-contained deployment (no .NET runtime installation required)

---

## High-Impact PLINQ Search Engine

XLSM TO JSON solves one of the biggest limitations of Excel-based workflow debugging: slow and unreliable searches over massive datasets containing hidden formatting characters, multiline values, arrays, and embedded logic.

Instead of relying on Excel’s cell rendering layer, Hammer Inspect preprocesses workbook data into normalized JSON structures and executes searches directly against raw semantic values. The search engine is powered by **Parallel LINQ (PLINQ)**, which distributes query execution across multiple CPU threads simultaneously.

### Technical Workflow

1. Excel rows are parsed and converted into clean in-memory JSON objects.
2. Search queries are executed using PLINQ parallel partitions.
3. The dataset is automatically split across available processor cores.
4. Each thread scans a subset of rows concurrently.
5. Matching rows are aggregated back into a unified result set.

This architecture eliminates the single-threaded bottlenecks commonly found in spreadsheet search operations and allows Hammer Inspect to scan more than **30,000 rows in milliseconds** while keeping the UI fully responsive.

### Example Search Flow

```csharp
var results = rows
    .AsParallel()
    .Where(row => row.JsonContent.Contains(searchText))
    .ToList();
```

By combining normalized JSON parsing with multithreaded PLINQ execution, Hammer Inspect delivers near-instant workflow inspection even on extremely large configuration files.

---

## Dependencies

### Runtime & Frameworks

- .NET 8
- WinUI 3
- Windows App SDK

### Libraries & Technologies

- Newtonsoft.Json
- Parallel LINQ (PLINQ)
- ClosedXML / Excel workbook parsing
- RichTextBlock rendering system
- GPU hardware acceleration via DirectWrite / WinUI compositor

---

## Deployment

- Architecture: Windows x64
- Deployment Type: Self-Contained
- Installer: Inno Setup

No additional .NET runtime or Windows App SDK installation is required.
