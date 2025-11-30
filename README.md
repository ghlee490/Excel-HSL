# Excel-HSL
A domain-specific language that solves optimization problems using the harmony search algorithm.

## 1. Files
- `Excel_HSL.xlam`: Excel add-in implementing the Harmony Search solver.
- `examples/`: four benchmark optimization models.
- `experiments/`: Python version code used for comparison (Jupyter notebook)

## 2. Installation (Excel Add-in)

1. Copy `Excel_HSL.xlam` to  
   `C:\Users\<YourUserName>\AppData\Roaming\Microsoft\AddIns`
2. In Excel: File → Options → Add-ins
3. At the bottom, set Manage: Excel Add-ins, then click Go…
4. In the Add-ins window:
If the add-in appears in the list → check the box next to it.
If it does not appear → click Browse…, navigate to the Add-ins folder, and select the .xlam file.
5. Click OK. The add-in is now loaded whenever Excel starts.

(Optional but Recommended)
6. Go to "File → Options → Customize Ribbon."
7. On the right side, choose a tab where you want to add the button
(e.g., Home, Data, or create a new custom tab).
8. Click New Group to create a dedicated section under the selected tab.
9. On the left side, set Choose commands from: → Macros.
10. Find your macro name (e.g., Excel_HSL).
11. Click Add >> to add it into the new group.
12. (Optional) Select the new macro in the right panel and click Rename…
 + Choose an icon, Set a friendly display name (e.g., Run Excel-HSL)
14. Click OK to apply.

Now you will see the macro button in your ribbon, and clicking it will run the add-in instantly.

## 3. Usage
An Excel-HSL model consists of three sections:
- [OBJ] — Defines the optimization objective (minimization or maximization) and the objective function expression.
- [VAR] — Defines the decision variables, including their names, bounds, and types.
- [ST] — Defines equality or inequality constraints.

A model must contain at least `[OBJ]` and `[VAR]`, while `[ST]` is optional if there are no constraints.
---
The `[OBJ]` tag specifies the optimization objective.  
The expression must begin with either:
- `min` — for minimization  
- `max` — for maximization  
Write the objective function after min or max.
ex) [OBJ] min x^2 + y^2
---
The `[VAR]` tag defines each decision variable on a separate line, in the following order:
  variable_name, minimum_value, maximum_value, type
- `type` can be:
  - `any` — real-valued variable  
  - `int` — integer-valued variable
ex) [VAR] x, 0, 10, any
---
Constraints are specified under the `[ST]` tag and may include equality or inequality  operators.  
During optimization, any candidate solution that violates a constraint is discarded by the Harmony Search engine.
ex) x+2*y <= 30

## 4. License
MIT License.
