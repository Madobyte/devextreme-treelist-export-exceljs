# DevExtreme TreeList - Export to Excel

### How to use
1. Add the following import statement:
> import { exportTreeList } from 'https://cdn.jsdelivr.net/gh/Madobyte/devextreme-treelist-export-exceljs@latest/excelExporter.js';
2. Add the ExcelJS and FileSaver packages.
3. Define the export button in the TreeList's toolbar.
4. Implement the button's onClick:
```
function exportToExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Employees");

    exportTreeList({
        component: treeList,
        worksheet,
    }).then(() => {
        workbook.xlsx.writeBuffer().then((buffer) => {
            saveAs(
                new Blob([buffer], { type: "application/octet-stream" }),
                "Employees.xlsx",
            );
        });
    });
}

```

### Sample
[CodePen](https://codepen.io/madobyte/pen/BaGbKOQ?editors=0010)
