(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // 每次加载新页面时都必须运行初始化函数。
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // 初始化通知机制并隐藏它
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // 如果未使用 Excel 2016，请使用回退逻辑。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("此示例将显示电子表格中选定单元格的值。");
                $('#button-text').text("显示!");
                $('#button-desc').text("显示所选内容");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("此示例将突出显示电子表格中选定单元格的最高值。");
            $('#button-text').text("突出显示!");
            $('#button-desc').text("突出显示最大数字。");
                
            loadSampleData();

            // 为突出显示按钮添加单击事件处理程序。
            $('#highlight-button').click(hightlightHighestValue);
         
            $("#run").click(run);
            $("#createWithNames").click(createWithNames);
            $("#setup").click(setup); 
            $("#setupPivot").click(setupPivot); 
            $("#createWithNames").click(addRow); 
            $("#newadd").click(deletePivot); 
            
            
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // 针对 Excel 对象模型运行批处理操作
        Excel.run(function (ctx) {
            // 为活动工作表创建代理对象
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // 将向电子表格写入示例数据的命令插入队列
            sheet.getRange("B3:D5").values = values;

            // 运行排队的命令，并返回承诺表示任务完成
            return ctx.sync();
        })
        .catch(errorHandler);
    }


    async function run() {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "red";
            range.load("address");

            await context.sync();

            console.log(`The range address was "${range.address}".`);
        });
    }


    async function createColumnClusteredChart() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("Sample");

            const salesTable = sheet.tables.getItem("SalesTable");

            const dataRange = salesTable.getDataBodyRange();

            let chart = sheet.charts.add("ColumnClustered", dataRange, "Auto");

            chart.setPosition("A9", "F20");
            chart.title.text = "Quarterly sales chart";
            chart.legend.position = "Right";
            chart.legend.format.fill.setSolidColor("white");
            chart.dataLabels.format.font.size = 15;
            chart.dataLabels.format.font.color = "black";
            let points = chart.series.getItemAt(0).points;
            points.getItemAt(0).format.fill.setSolidColor("pink");
            points.getItemAt(1).format.fill.setSolidColor("indigo");

            await context.sync();
        });
    }

    async function run1() {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "blue";
            range.load("address");

            await context.sync();

            console.log(`The range address was "${range.address}".`);
        });
    }
    async function createWithNames() {
        await Excel.run(async (context) => {
            const rangeToAnalyze = context.workbook.worksheets.getItem("Data").getRange("A1:E21");
            const rangeToPlacePivot = context.workbook.worksheets.getItem("Pivot").getRange("A2");
            context.workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);

            await context.sync();
        });
    }
    async function setupPivot() {
        await Excel.run(async (context) => {
            context.workbook.worksheets.getItemOrNullObject("Data").delete();
            const dataSheet = context.workbook.worksheets.add("Data");
            context.workbook.worksheets.getItemOrNullObject("Pivot").delete();
            const pivotSheet = context.workbook.worksheets.add("Pivot");

            const data = [
                ["Farm", "Type", "Classification", "Crates Sold at Farm", "Crates Sold Wholesale"],
                ["A Farms", "Lime", "Organic", 300, 2000],
                ["A Farms", "Lemon", "Organic", 250, 1800],
                ["A Farms", "Orange", "Organic", 200, 2200],
                ["B Farms", "Lime", "Conventional", 80, 1000],
                ["B Farms", "Lemon", "Conventional", 75, 1230],
                ["B Farms", "Orange", "Conventional", 25, 800],
                ["B Farms", "Orange", "Organic", 20, 500],
                ["B Farms", "Lemon", "Organic", 10, 770],
                ["B Farms", "Kiwi", "Conventional", 30, 300],
                ["B Farms", "Lime", "Organic", 50, 400],
                ["C Farms", "Apple", "Organic", 275, 220],
                ["C Farms", "Kiwi", "Organic", 200, 120],
                ["D Farms", "Apple", "Conventional", 100, 3000],
                ["D Farms", "Apple", "Organic", 80, 2800],
                ["E Farms", "Lime", "Conventional", 160, 2700],
                ["E Farms", "Orange", "Conventional", 180, 2000],
                ["E Farms", "Apple", "Conventional", 245, 2200],
                ["E Farms", "Kiwi", "Conventional", 200, 1500],
                ["F Farms", "Kiwi", "Organic", 100, 150],
                ["F Farms", "Lemon", "Conventional", 150, 270]
            ];

            const range = dataSheet.getRange("A1:E21");
            range.values = data;
            range.format.autofitColumns();

            pivotSheet.activate();

            await context.sync();
        });
    }


    async function setupPivot1() {
        await Excel.run(async (context) => {
            context.workbook.worksheets.getItemOrNullObject("Data").delete();
            const dataSheet = context.workbook.worksheets.add("Data");
            context.workbook.worksheets.getItemOrNullObject("Pivot").delete();
            const pivotSheet = context.workbook.worksheets.add("Pivot");
            //
          const apiUrl = 'https://localhost:44324/Home/GetData'; 
            //const apiUrl = 'https://localhost:56932/Home/GetData?name=tom';
            const response = await fetch(apiUrl); // 发送请求获取数据 
            const data = await response.json(); // 将响应数据解析为 JSON 格式



            /*
            const data = [
                ["Farm", "Type", "Classification", "Crates Sold at Farm", "Crates Sold Wholesale"],
                ["A Farms", "Lime", "Organic", 300, 2000],
                ["A Farms", "Lemon", "Organic", 250, 1800],
                ["A Farms", "Orange", "Organic", 200, 2200],
                ["B Farms", "Lime", "Conventional", 80, 1000],
                ["B Farms", "Lemon", "Conventional", 75, 1230],
                ["B Farms", "Orange", "Conventional", 25, 800],
                ["B Farms", "Orange", "Organic", 20, 500],
                ["B Farms", "Lemon", "Organic", 10, 770],
                ["B Farms", "Kiwi", "Conventional", 30, 300],
                ["B Farms", "Lime", "Organic", 50, 400],
                ["C Farms", "Apple", "Organic", 275, 220],
                ["C Farms", "Kiwi", "Organic", 200, 120],
                ["D Farms", "Apple", "Conventional", 100, 3000],
                ["D Farms", "Apple", "Organic", 80, 2800],
                ["E Farms", "Lime", "Conventional", 160, 2700],
                ["E Farms", "Orange", "Conventional", 180, 2000],
                ["E Farms", "Apple", "Conventional", 245, 2200],
                ["E Farms", "Kiwi", "Conventional", 200, 1500],
                ["F Farms", "Kiwi", "Organic", 100, 150],
                ["F Farms", "Lemon", "Conventional", 150, 270]
            ];
           */
            const range = dataSheet.getRange("A1:E21");
            range.values = data;
            range.format.autofitColumns();

            pivotSheet.activate();

            await context.sync();
        });
    }


    async function addRow() {
        await Excel.run(async (context) => {
            const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

            // Check if the PivotTable already has rows.
            const farmRow = pivotTable.rowHierarchies.getItemOrNullObject("Farm");
            const typeRow = pivotTable.rowHierarchies.getItemOrNullObject("Type");
            const classificationRow = pivotTable.rowHierarchies.getItemOrNullObject("Classification");
            pivotTable.rowHierarchies.load();
            await context.sync();

            if (farmRow.isNullObject) {
                pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
            } else if (typeRow.isNullObject) {
                pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
            } else if (classificationRow.isNullObject) {
                pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
            }

            await context.sync();
        });
    }

    async function setup() {
        await Excel.run(async (context) => {
            context.workbook.worksheets.getItemOrNullObject("Sample").delete();
            const sheet = context.workbook.worksheets.add("Sample");

            let expensesTable = sheet.tables.add("A1:E1", true);
            expensesTable.name = "SalesTable";
            expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

            expensesTable.rows.add(null, [
                ["Frames", 5000, 7000, 6544, 4377],
                ["Saddles", 400, 323, 276, 651],
                ["Brake levers", 12000, 8766, 8456, 9812],
                ["Chains", 1550, 1088, 692, 853],
                ["Mirrors", 225, 600, 923, 544],
                ["Spokes", 6005, 7634, 4589, 8765]
            ]);

            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();

            sheet.activate();
            await context.sync();
        });
    }
    function hightlightHighestValue() {
        // 针对 Excel 对象模型运行批处理操作
        Excel.run(function (ctx) {
            // 创建选定范围的代理对象，并加载其属性
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // 运行排队的命令，并返回承诺表示任务完成
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // 找到要突出显示的单元格
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // 突出显示该单元格
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('选定的文本为:', '"' + result.value + '"');
                } else {
                    showNotification('错误', result.error.message);
                }
            });
    }

    // 处理错误的帮助程序函数
    function errorHandler(error) {
        // 请务必捕获 Excel.run 执行过程中出现的所有累积错误
        showNotification("错误", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 用于显示通知的帮助程序函数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }


    async function deletePivot() {
        await Excel.run(async (context) => {
            context.workbook.worksheets
                .getItem("Pivot")
                .pivotTables.getItem("Farm Sales")
                .delete();

            // Also clean up the extra data from getCrateTotal().
            context.workbook.worksheets
                .getActiveWorksheet()
                .getRange("B27:C27")
                .delete(Excel.DeleteShiftDirection.up);
            await context.sync();
        });
    }
})();
