var summaryJsonData = {{summaryJson}};
// var combinedJsonData = {{combinedJson}};
console.log(summaryJsonData);
// console.log(combinedJsonData);
const typeList =[
      "RetailMargin",
      "Revenue",
      "Network",
      "Capacity",
      "WholesaleEnergy",
      "MarketFees",
      "ESS",
      "LGC",
      "STC",
      "Commission"
  ]
const achievedMargin2024 = transformDataForPieChart(summaryJsonData, "Sum of Achieved Margin FY2024")
const dataFy2024 =  transformDataForPieChart(summaryJsonData, "FY2024")
const dataFy2025 =  transformDataForPieChart(summaryJsonData, "FY2025")
const dataFy2026 = transformDataForPieChart(summaryJsonData, "FY2026")

const transformDataForPieChart = (data, marginType) => {
    // Find the entry for the specified typesofmargin
    const marginData = data.find(entry => entry.typesofmargin.includes(marginType));
  
    if (!marginData) {
      console.error("No data found for the specified typesofmargin.");
      return [];
    }
  
    // Convert totals to array of objects with 'type' and 'total'
    return Object.keys(marginData.totals).map(key => ({
      type: key,
      total: parseFloat(marginData.totals[key])
    }));
};
  

const animatedPieChart = (data, chartId) => {
  am4core.useTheme(am4themes_animated);
  var chart = am4core.create(chartId, am4charts.PieChart3D);

  chart.data = data;
  chart.innerRadius = am4core.percent(40);

  var pieSeries = chart.series.push(new am4charts.PieSeries3D());
  pieSeries.dataFields.value = "total";
  pieSeries.dataFields.category = "type";

  pieSeries.labels.template.disabled = chartId != "acap";
  pieSeries.ticks.template.disabled = chartId != "acap";

  // Add a legend
  chart.legend = new am4charts.Legend();
  chart.legend.position = "right";
};

const xyChart = (data, chartId, seriesName) => {
  am4core.useTheme(am4themes_animated);
  var chart = am4core.create(chartId, am4charts.XYChart);

  chart.data = data;
  var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
  categoryAxis.dataFields.category = "type";
  categoryAxis.title.text = "Margins";

  // Set minimum grid distance to a smaller value to fit more labels
  categoryAxis.renderer.minGridDistance = 20; // Adjust this value as needed

  // Disable automatic label hiding
  categoryAxis.renderer.labels.template.adapter.add("dy", function (dy, target) {
  return 0; // Keep all labels visible
  });

  // Optionally, rotate labels if there are too many
  categoryAxis.renderer.labels.template.rotation = 45; // Rotate labels 45 degrees
  categoryAxis.renderer.labels.template.horizontalCenter = "right";
  categoryAxis.renderer.labels.template.verticalCenter = "middle";
  categoryAxis.renderer.labels.template.fontSize = 12;


  var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());
  valueAxis.title.text = "Value";

  var series = chart.series.push(new am4charts.ColumnSeries());
  series.dataFields.valueY = "total";
  series.dataFields.categoryX = "type";
  series.name = seriesName
  series.stacked = true;

  series.columns.template.tooltipText = "{categoryX}: [bold]{valueY}[/]";
  series.columns.template.fontSize = 14; // Set font size for series labels

  series.tooltip.fontSize = 12; // Customize tooltip font size

  chart.yAxes.getIndex(0).renderer.labels.template.fontSize = 12; // Y-axis labels
  chart.xAxes.getIndex(0).renderer.labels.template.fontSize = 12;

  chart.cursor = new am4charts.XYCursor();
};

const populateTable1Data = (data) => {
  const headerRow = document.querySelector('#table1 #headerRow');
  const dataRow = document.querySelector('#table1 #dataRow');
  
  data.forEach(item => {
      const th = document.createElement('th');
      th.textContent = item.type.replace(/([A-Z][a-z]+|[A-Z]+(?![a-z]))/g, ' $1').trim();
      headerRow.appendChild(th);
  });

  // Populate data dynamically
  data.forEach(item => {
      const td = document.createElement('td');
      td.textContent = parseFloat(item.total).toFixed(2)
      dataRow.appendChild(td);
  });
}
const clusteredChart = (data, chartId) => {
    am4core.useTheme(am4themes_animated);
    var chart = am4core.create(chartId, am4charts.XYChart);
    chart.colors.step = 2;

    chart.legend = new am4charts.Legend();
    chart.legend.position = 'top';
    chart.legend.paddingBottom = 20;
    chart.legend.labels.template.maxWidth = 95;

    var xAxis = chart.xAxes.push(new am4charts.CategoryAxis());
    xAxis.dataFields.category = 'type';
    xAxis.renderer.cellStartLocation = 0.1;
    xAxis.renderer.cellEndLocation = 0.9;
    xAxis.renderer.grid.template.location = 0;

    function getExtremes(data) {
        const initialMax = -Infinity;
        const initialMin = Infinity;
        
        const extremes = data.reduce((acc, item) => {
            const value = parseFloat(item.total);

            // Update max and min values
            if (value > acc.max) {
                acc.max = value;
            }
            if (value < acc.min) {
                acc.min = value;
            }

            return acc;
        }, { max: initialMax, min: initialMin });

        return extremes;
    }

    // Example usage
    const { max: highestValue, min: lowestValue } = getExtremes(data);


    var yAxis = chart.yAxes.push(new am4charts.ValueAxis()); // By default, this handles negative values too
    yAxis.min = lowestValue; // Adjust as needed to ensure negative values are visible
    yAxis.max = highestValue;
    yAxis.renderer.minGridDistance = 20; // Adjust spacing for clarity if needed

    function createSeries(value, name) {
        var series = chart.series.push(new am4charts.ColumnSeries());
        series.dataFields.valueY = value;
        series.dataFields.categoryX = 'type';
        series.name = name;
        series.columns.template.tooltipText = "{categoryX}: [bold]{valueY}[/]";
        series.columns.template.fontSize = 14; // Set font size for series labels

        series.events.on("hidden", arrangeColumns);
        series.events.on("shown", arrangeColumns);

        var bullet = series.bullets.push(new am4charts.LabelBullet());
        bullet.interactionsEnabled = true;
        bullet.dy = 30;
        bullet.label.fill = am4core.color('#ffffff');

        return series;
    }

    // Map data to the chart
    const categories = [...new Set(data.map(item => item.type))]; // Unique types
    const seriesNames = [...new Set(data.map(item => item.typesofmargin))]; // Unique margin types

    // Prepare data structure for the chart
    const chartData = categories.map(category => {
        const dataPoint = { type: category };
        seriesNames.forEach(name => {
            const entry = data.find(d => d.type === category && d.typesofmargin === name);
            dataPoint[name] = entry ? parseFloat(entry.total) : 0;
        });
        return dataPoint;
    });

    chart.data = chartData;

    // Create series for each typesofmargin
    seriesNames.forEach(name => {
        createSeries(name, name);
    });

    function arrangeColumns() {
        var series = chart.series.getIndex(0);

        var w = 1 - xAxis.renderer.cellStartLocation - (1 - xAxis.renderer.cellEndLocation);
        if (series.dataItems.length > 1) {
            var x0 = xAxis.getX(series.dataItems.getIndex(0), "categoryX");
            var x1 = xAxis.getX(series.dataItems.getIndex(1), "categoryX");
            var delta = ((x1 - x0) / chart.series.length) * w;
            if (am4core.isNumber(delta)) {
                var middle = chart.series.length / 2;

                var newIndex = 0;
                chart.series.each(function(series) {
                    if (!series.isHidden && !series.isHiding) {
                        series.dummyData = newIndex;
                        newIndex++;
                    }
                    else {
                        series.dummyData = chart.series.indexOf(series);
                    }
                })
                var visibleCount = newIndex;
                var newMiddle = visibleCount / 2;

                chart.series.each(function(series) {
                    var trueIndex = chart.series.indexOf(series);
                    var newIndex = series.dummyData;

                    var dx = (newIndex - trueIndex + middle - newMiddle) * delta;

                    series.animate({ property: "dx", to: dx }, series.interpolationDuration, series.interpolationEasing);
                    series.bulletsContainer.animate({ property: "dx", to: dx }, series.interpolationDuration, series.interpolationEasing);
                })
            }
        }
    }
}
const populateTable2Data = (data, id) => {
    const headerRow = document.querySelector(`#${id} #headerRow`);
    const tbody = document.querySelector(`#${id} tbody`);

    // Create a set to hold unique types of margins
    const marginTypes = new Set();
    const typesOfMargins = new Set();

    // Populate the sets with unique types and margin types from data
    data.forEach(item => {
        marginTypes.add(item.type);
        typesOfMargins.add(item.typesofmargin);
    });

    // Create table headers
    headerRow.innerHTML = '<th></th>' + Array.from(marginTypes).map(type => 
        `<th>${type.replace(/([A-Z][a-z]+|[A-Z]+(?![a-z]))/g, ' $1').trim()}</th>`
    ).join('');

    // Create a map to hold totals for each margin type and each type
    const totalsMap = {};

    // Initialize the totals map
    typesOfMargins.forEach(marginType => {
        totalsMap[marginType] = {};
        marginTypes.forEach(type => {
            totalsMap[marginType][type] = 0;
        });
    });

    // Populate the map with totals from data
    data.forEach(item => {
        if (totalsMap[item.typesofmargin]) {
            totalsMap[item.typesofmargin][item.type] += parseFloat(item.total);
        }
    });

    // Populate data rows
    Object.keys(totalsMap).forEach(marginType => {
        const row = document.createElement('tr');
        const marginTypeCell = document.createElement('td');
        marginTypeCell.textContent = `Sum of ${marginType}`;
        row.appendChild(marginTypeCell);

        marginTypes.forEach(type => {
            const td = document.createElement('td');
            td.textContent = totalsMap[marginType][type].toFixed(2);
            row.appendChild(td);
        });

        tbody.appendChild(row);
    });
};

const totalCost = (data) => {
    // Initialize an empty object to hold totals for each type
    const totals = {};

    // Populate totals with sums for each type
    data.forEach(item => {
        if (!totals[item.type]) {
            totals[item.type] = 0;
        }
        totals[item.type] += parseFloat(item.total);
    });

    // Convert the totals object to an array of objects
    return Object.keys(totals).map(type => ({
        type: type,
        total: totals[type].toFixed(2) // Format to two decimal places
    }));
};
const totalCostByMargin = (data) => {
    const totalsByMargin = {};

    data.forEach(item => {
        const { typesofmargin, type, total } = item;
        
        if (!totalsByMargin[typesofmargin]) {
            totalsByMargin[typesofmargin] = {};
        }
        
        if (!totalsByMargin[typesofmargin][type]) {
            totalsByMargin[typesofmargin][type] = 0;
        }
        totalsByMargin[typesofmargin][type] += parseFloat(total);
    });

    // Convert the results into an array of objects
    return Object.keys(totalsByMargin).map(marginType => ({
        typesofmargin: marginType,
        totals: Object.keys(totalsByMargin[marginType])
            .map(type => ({
                type: type,
                total: totalsByMargin[marginType][type].toFixed(2) // Format to two decimal places
            }))
            .sort((a, b) => a.type.localeCompare(b.type)) // Sort totals by type
    }));
};
const fyActualMarginVsPOE = (dataArr, chartId) => {
    am4core.useTheme(am4themes_animated);

    // Data
    const data = dataArr;

    // Transform data
    const transformedData = [];
    data.forEach(margin => {
        margin.totals.forEach(total => {
            transformedData.push({
                typesofmargin: margin.typesofmargin,
                type: total.type,
                total: parseFloat(total.total)
            });
        });
    });
    console.log(transformedData);
    
    // Create chart instance
    var chart = am4core.create(chartId, am4charts.XYChart);

    // Add data
    chart.data = transformedData;

    // Create X Axis
    var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
    categoryAxis.dataFields.category = "type";
    categoryAxis.renderer.labels.template.fontSize = 14;
    categoryAxis.renderer.grid.template.location = 0;
    categoryAxis.renderer.minGridDistance = 20;
    categoryAxis.title.text = "Type";

    // Create Y Axis
    var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());
    valueAxis.renderer.inversed = false; // Set to true if you want a reversed value axis
    valueAxis.renderer.minWidth = 35;
    valueAxis.title.text = "Total";

    // Create series for each typesofmargin
    data.forEach(margin => {
        var series = chart.series.push(new am4charts.LineSeries());
        series.dataFields.valueY = "total";
        series.dataFields.categoryX = "type";
        series.dataFields.name = "typesofmargin";
        series.name = margin.typesofmargin;
        series.strokeWidth = 2;
        series.tooltipText = "{categoryX}: [bold]{valueY}[/]";

        // Add bullets to the series
        var bullet = series.bullets.push(new am4charts.CircleBullet());
        bullet.circle.fill = am4core.color("#FF5733"); // Bullet color (you can use a function to set dynamic colors)
        bullet.circle.stroke = am4core.color("#fff"); // Bullet border color
        bullet.circle.strokeWidth = 2; // Bullet border width
        bullet.circle.radius = 6; // Bullet size
    });

    // Add cursor
    chart.cursor = new am4charts.XYCursor();
    chart.cursor.behavior = "zoomY";

    // Add legend
    chart.legend = new am4charts.Legend();

    // Add scrollbar
    chart.scrollbarX = new am4core.Scrollbar();
}

xyChart(achievedMargin2024, "hpfy24a", "Historical Performance FY2024 Actuals");
animatedPieChart(achievedMargin2024, "pfify24");
populateTable1Data(achievedMargin2024);
clusteredChart(dataFy2025, "fy25avspoe");
populateTable2Data(dataFy2025, 'table2');
clusteredChart(dataFy2026, "fy26avspoe");
populateTable2Data(dataFy2026, 'table3');
animatedPieChart(totalCost(summaryJsonData), "acap");

// Example usage
console.log(totalCostByMargin(summaryJsonData));
fyActualMarginVsPOE(totalCostByMargin(summaryJsonData), "fyamvspoe");