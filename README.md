
<img width="1038" height="546" alt="image" src="https://github.com/user-attachments/assets/3d45367f-6036-482d-97ad-2c659f6640e1" />


# InlineCppVarDbg

**InlineCppVarDbg** is a lightweight Visual Studio extension that adds inline value hints for C++ debugging, so you can inspect important values directly in the editor without constantly switching to the **Locals** or **Watch** windows.

Designed for **Visual Studio 2022** and **Visual Studio 2026** (`17.x` / `18.x`), it displays values while execution is paused in break mode, with a strong focus on usability, clarity, and step-performance safety.

## Marketplace

[Install from Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=podeba.InlineCppDbg)

## Highlights

- Supports both **inline** and **end-of-line** value display modes
- Configurable visible range for previous lines, scoped to the **current function**
- Always shows **function parameter values**
- Convenient one-click toolbar toggles for:
  - Enable / disable
  - Display mode
  - Number format (`Dec` / `Hex` / `Bin`)
- Clean rendering for:
  - Arrays
  - Enums
  - Null pointers
  - Uninitialized values
- Highlights changed values with a **step-based fade effect**
- Customizable **chip colors** and **font size**
- Automatically ignores inactive `#ifdef` regions
- Includes a **performance fail-safe** with a strict evaluation budget and graceful fallback behavior to avoid slowing down step operations

## Why InlineCppVarDbg?

When debugging C++, constantly checking variable values in separate tool windows can interrupt your flow. InlineCppVarDbg keeps relevant values close to the code, helping you understand program state faster while keeping stepping responsive and predictable.
