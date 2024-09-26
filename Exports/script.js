const transformDataForChart = (data, marginType) => {
  const marginData = data.find((entry) =>
    entry.typesofmargin.includes(marginType)
  );
  if (!marginData) {
    console.error("No data found for the specified typesofmargin.");
    return [];
  }

  return Object.keys(marginData.totals)
    .filter((key) => key.toLowerCase() !== "total")
    .map((key) => ({
      type: marginData.labels[key] || key,
      total: parseFloat(marginData.totals[key]),
    }));
};
const transformDataForCluster = (data, marginType) => {
  var marginData = [];
  if (marginType == null || marginType == "") {
    marginData = data;
  } else {
    marginData = data.filter((entry) =>
      entry.typesofmargin.includes(marginType)
    );
  }

  const combineData = (data) => {
    const combinedArray = [];

    data.forEach((entry) => {
      const { typesofmargin, totals, labels } = entry;

      Object.keys(totals).forEach((key) => {
        if (key.toLowerCase() !== "total") {
          combinedArray.push({
            margin: typesofmargin,
            type: labels[key] || key,
            total: parseFloat(totals[key]),
          });
        }
      });
    });

    return combinedArray;
  };

  if (!marginData) {
    console.error("No data found for the specified typesofmargin.");
    return [];
  }

  return combineData(marginData);
};

const animatedPieChart = (data, chartId) => {
  am4core.useTheme(am4themes_animated);
  var chart = am4core.create(chartId, am4charts.PieChart3D);

  chart.data = data;
  chart.innerRadius = am4core.percent(40);

  var pieSeries = chart.series.push(new am4charts.PieSeries3D());
  pieSeries.dataFields.value = "total";
  pieSeries.dataFields.category = "type";

  if (chartId != "acap") {
    pieSeries.alignLabels = false;
    pieSeries.labels.template.adapter.add("text", function (text, target) {
      if (
        target.dataItem &&
        Math.abs(target.dataItem.values.value.percent) < 1
      ) {
        return ""; // Return empty string to hide the label
      }
      return text;
    });
  }

  chart.legend = new am4charts.Legend();
  chart.legend.position = "right";
  chart.exporting.menu = new am4core.ExportMenu();
};

const xyChart = (data, chartId, seriesName) => {
  am4core.useTheme(am4themes_animated);
  var chart = am4core.create(chartId, am4charts.XYChart);

  chart.data = data;
  var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
  categoryAxis.dataFields.category = "type";
  categoryAxis.title.text = "Margins";

  categoryAxis.renderer.minGridDistance = 20;

  categoryAxis.renderer.labels.template.adapter.add(
    "dy",
    function (dy, target) {
      return 0;
    }
  );

  categoryAxis.renderer.labels.template.rotation = 45;
  categoryAxis.renderer.labels.template.horizontalCenter = "right";
  categoryAxis.renderer.labels.template.verticalCenter = "middle";
  categoryAxis.renderer.labels.template.fontSize = 12;

  var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());
  valueAxis.title.text = "Value";

  var series = chart.series.push(new am4charts.ColumnSeries());
  series.dataFields.valueY = "total";
  series.dataFields.categoryX = "type";
  series.name = seriesName;
  series.stacked = true;

  series.columns.template.tooltipText = "{categoryX}: [bold]{valueY}[/]";
  series.columns.template.fontSize = 14;

  series.tooltip.fontSize = 12;

  chart.yAxes.getIndex(0).renderer.labels.template.fontSize = 12;
  chart.xAxes.getIndex(0).renderer.labels.template.fontSize = 12;

  chart.cursor = new am4charts.XYCursor();
  chart.exporting.menu = new am4core.ExportMenu();
};

const populateTable1Data = (data) => {
  const headerRow = document.querySelector("#table1 #headerRow");
  const dataRow = document.querySelector("#table1 #dataRow");

  data.forEach((item) => {
    const th = document.createElement("th");
    th.textContent = item.type
      .replace(/([A-Z][a-z]+|[A-Z]+(?![a-z]))/g, " $1")
      .trim();
    headerRow.appendChild(th);
  });

  data.forEach((item) => {
    const td = document.createElement("td");
    td.textContent = parseFloat(item.total).toFixed(2);
    dataRow.appendChild(td);
  });
};

const clusteredChart = (data, chartId) => {
  am4core.useTheme(am4themes_animated);
  var chart = am4core.create(chartId, am4charts.XYChart);
  chart.colors.step = 2;

  chart.cursor = new am4charts.XYCursor();
  chart.cursor.behavior = "zoomXY";

  chart.legend = new am4charts.Legend();
  chart.legend.position = "top";
  chart.legend.paddingBottom = 20;
  chart.legend.labels.template.maxWidth = 95;

  var xAxis = chart.xAxes.push(new am4charts.CategoryAxis());
  xAxis.dataFields.category = "type";
  xAxis.renderer.cellStartLocation = 0.1;
  xAxis.renderer.cellEndLocation = 0.9;
  xAxis.renderer.grid.template.location = 0;

  function getExtremes(data) {
    const initialMax = -Infinity;
    const initialMin = Infinity;

    const extremes = data.reduce(
      (acc, item) => {
        const value = parseFloat(item.total);

        if (value > acc.max) {
          acc.max = value;
        }
        if (value < acc.min) {
          acc.min = value;
        }

        return acc;
      },
      { max: initialMax, min: initialMin }
    );

    return extremes;
  }

  const { max: highestValue, min: lowestValue } = getExtremes(data);

  var yAxis = chart.yAxes.push(new am4charts.ValueAxis());
  yAxis.min = chartId == "rmpnmibase" ? -50000 : lowestValue;
  yAxis.max = chartId == "rmpnmibase" ? 50000 : highestValue;
  yAxis.renderer.minGridDistance = 20;

  function createSeries(value, name) {
    var series = chart.series.push(new am4charts.ColumnSeries());
    series.dataFields.valueY = value;
    series.dataFields.categoryX = "type";
    series.name = name;
    if (chartId == "rmpnmibase") {
      series.columns.template.tooltipText = "{name}: [bold]{valueY}[/]";
    } else {
      series.columns.template.tooltipText = "{categoryX}: [bold]{valueY}[/]";
    }

    series.columns.template.fontSize = 14;

    series.events.on("hidden", arrangeColumns);
    series.events.on("shown", arrangeColumns);

    var bullet = series.bullets.push(new am4charts.LabelBullet());
    bullet.interactionsEnabled = true;
    bullet.dy = 30;
    bullet.label.fill = am4core.color("#ffffff");

    return series;
  }

  const categories = [...new Set(data.map((item) => item.type))];
  const seriesNames = [...new Set(data.map((item) => item.margin))];

  const chartData = categories.map((category) => {
    const dataPoint = { type: category };
    seriesNames.forEach((name) => {
      const entry = data.find((d) => d.type === category && d.margin === name);
      dataPoint[name] = entry ? parseFloat(entry.total) : 0;
    });
    return dataPoint;
  });

  chart.data = chartData;

  seriesNames.forEach((name) => {
    createSeries(name, name);
  });

  function arrangeColumns() {
    var series = chart.series.getIndex(0);

    var w =
      1 -
      xAxis.renderer.cellStartLocation -
      (1 - xAxis.renderer.cellEndLocation);
    if (series.dataItems.length > 1) {
      var x0 = xAxis.getX(series.dataItems.getIndex(0), "categoryX");
      var x1 = xAxis.getX(series.dataItems.getIndex(1), "categoryX");
      var delta = ((x1 - x0) / chart.series.length) * w;
      if (am4core.isNumber(delta)) {
        var middle = chart.series.length / 2;

        var newIndex = 0;
        chart.series.each(function (series) {
          if (!series.isHidden && !series.isHiding) {
            series.dummyData = newIndex;
            newIndex++;
          } else {
            series.dummyData = chart.series.indexOf(series);
          }
        });
        var visibleCount = newIndex;
        var newMiddle = visibleCount / 2;

        chart.series.each(function (series) {
          var trueIndex = chart.series.indexOf(series);
          var newIndex = series.dummyData;

          var dx = (newIndex - trueIndex + middle - newMiddle) * delta;

          series.animate(
            { property: "dx", to: dx },
            series.interpolationDuration,
            series.interpolationEasing
          );
          series.bulletsContainer.animate(
            { property: "dx", to: dx },
            series.interpolationDuration,
            series.interpolationEasing
          );
        });
      }
    }
  }

  chart.exporting.menu = new am4core.ExportMenu();
};

const populateTable2Data = (data, id) => {
  const headerRow = document.querySelector("#" + id + " #headerRow");
  const tbody = document.querySelector("#" + id + " tbody");

  const marginTypes = new Set();
  const typesOfMargins = new Set();

  data.forEach((item) => {
    marginTypes.add(item.type);
    typesOfMargins.add(item.margin);
  });

  headerRow.innerHTML =
    "<th></th>" +
    Array.from(marginTypes)
      .map(
        (type) =>
          "<th>" +
          type.replace(/([A-Z][a-z]+|[A-Z]+(?![a-z]))/g, " $1").trim() +
          "</th>"
      )
      .join("");

  const totalsMap = {};

  typesOfMargins.forEach((marginType) => {
    totalsMap[marginType] = {};
    marginTypes.forEach((type) => {
      totalsMap[marginType][type] = 0;
    });
  });

  data.forEach((item) => {
    if (totalsMap[item.margin]) {
      totalsMap[item.margin][item.type] += parseFloat(item.total);
    }
  });

  Object.keys(totalsMap).forEach((marginType) => {
    const row = document.createElement("tr");
    const marginTypeCell = document.createElement("td");
    marginTypeCell.textContent = marginType;
    row.appendChild(marginTypeCell);

    marginTypes.forEach((type) => {
      const td = document.createElement("td");
      td.textContent = totalsMap[marginType][type].toFixed(2);
      row.appendChild(td);
    });

    tbody.appendChild(row);
  });
};

const populateTable3Data = (data, id) => {
  const headerRow = document.querySelector("#" + id + " #headerRow");
  const tbody = document.querySelector("#" + id + " tbody");

  headerRow.innerHTML = "";
  tbody.innerHTML = "";

  headerRow.innerHTML = "<th>Margin</th><th>Total</th>";

  const totalsMap = {};

  data.forEach((item) => {
    if (!totalsMap[item.type]) {
      totalsMap[item.type] = 0;
    }
    totalsMap[item.type] += parseFloat(item.total);
  });

  Object.keys(totalsMap).forEach((type) => {
    const row = document.createElement("tr");

    const typeCell = document.createElement("td");
    typeCell.textContent = type;
    row.appendChild(typeCell);

    const totalCell = document.createElement("td");
    totalCell.textContent = totalsMap[type].toFixed(2);
    row.appendChild(totalCell);

    tbody.appendChild(row);
  });
};

const getUniqueValue = (data, element) => {
  const uniqueTypes = [
    ...new Set(
      data
        .map((item) => item.type)
        .filter((type) => type && type.trim() !== "(blank)")
    ),
  ];
  const selectAllOption = document.createElement("option");
  selectAllOption.value = "selectAll";
  selectAllOption.text = "Select All";
  element.insertBefore(selectAllOption, filter3Nmi.firstChild);

  uniqueTypes.forEach((type) => {
    const option = document.createElement("option");
    option.value = type;
    option.text = type;
    element.appendChild(option);
  });
};

const setFilterData = (data, element, key) => {
  const uniqueTypes = [
    ...new Set(
      data
        .map((item) => item[key])
        .filter((type) => type && type.trim() !== "(blank)")
    ),
  ];
  const selectAllOption = document.createElement("option");
  selectAllOption.value = "selectAll";
  selectAllOption.text = "Select All";
  element.insertBefore(selectAllOption, element.firstChild);

  uniqueTypes.forEach((type) => {
    const option = document.createElement("option");
    option.value = type;
    option.text = type;
    element.appendChild(option);
  });
};

const getYearFilter = (data, element, key) => {
  let uniqueTypes = [];

  if (key === "margin") {
    // Accessing the margin property correctly
    uniqueTypes = data[0].data.map((item) => item.margin);
  } else if (key == "nmi") {
    uniqueTypes = data.map((item) => item.margin);
    console.log(uniqueTypes);
  } else {
    // Ensuring margin and slice are accessed correctly
    uniqueTypes = [
      ...new Set(
        data
          .flatMap(
            (obj) =>
              obj.data
                .map((item) => item.margin.match(/FY\d{4}/)) // Match 'FY' followed by 4 digits
                .filter(Boolean) // Remove null values if no match is found
          )
          .flat() // Flatten the matched arrays
      ),
    ];
  }

  // Create and insert "Select All" option
  const selectAllOption = document.createElement("option");
  selectAllOption.value = "selectAll";
  selectAllOption.text = "Select All";
  element.insertBefore(selectAllOption, element.firstChild);

  // Append the unique types to the element
  uniqueTypes.forEach((type) => {
    const option = document.createElement("option");
    option.value = type;
    option.text = type;
    element.appendChild(option);
  });
};

const totalCost = (data) => {
  const totals = {};

  data.forEach((item) => {
    if (!totals[item.type]) {
      totals[item.type] = 0;
    }
    totals[item.type] += parseFloat(item.total);
  });

  return Object.keys(totals).map((type) => ({
    type: type,
    total: totals[type].toFixed(2),
  }));
};

const fyActualMarginVsPOE = (dataArr, chartId) => {
  am4core.useTheme(am4themes_animated);

  var chart = am4core.create(chartId, am4charts.XYChart);

  const transformedData = [];

  dataArr.forEach((margin) => {
    margin.type.forEach((total) => {
      let existingEntry = transformedData.find(
        (entry) => entry.type === total.name
      );
      if (!existingEntry) {
        existingEntry = { type: total.name };
        transformedData.push(existingEntry);
      }

      existingEntry[margin.margin] = total.value;
    });
  });

  chart.data = transformedData;

  var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
  categoryAxis.dataFields.category = "type";
  categoryAxis.renderer.labels.template.fontSize = 14;
  categoryAxis.renderer.grid.template.location = 0;
  categoryAxis.renderer.minGridDistance = 20;
  categoryAxis.title.text = "Type";

  var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());
  valueAxis.renderer.minWidth = 35;
  valueAxis.title.text = "Value";

  dataArr.forEach((margin) => {
    var series = chart.series.push(new am4charts.LineSeries());
    series.dataFields.valueY = margin.margin;
    series.dataFields.categoryX = "type";
    series.name = margin.margin;
    series.strokeWidth = 2;
    series.tooltipText = "{name}: [bold]{valueY}[/]";

    var bullet = series.bullets.push(new am4charts.CircleBullet());
    bullet.circle.fill = am4core.color("#FF5733");
    bullet.circle.stroke = am4core.color("#fff");
    bullet.circle.strokeWidth = 2;
    bullet.circle.radius = 6;
  });

  chart.cursor = new am4charts.XYCursor();
  chart.cursor.behavior = "zoomY";

  chart.legend = new am4charts.Legend();

  chart.scrollbarX = new am4core.Scrollbar();
  chart.exporting.menu = new am4core.ExportMenu();
};

const nmiClusteredChart = (data, chartId) => {
  loadingElement.style.display = "block";
  if (nmiChartVar) {
    nmiChartVar.dispose();
  }
  am4core.useTheme(am4themes_animated);
  nmiChartVar = am4core.create(chartId, am4charts.XYChart);
  nmiChartVar.colors.step = 2;

  nmiChartVar.cursor = new am4charts.XYCursor();
  nmiChartVar.cursor.behavior = "zoomXY";

  nmiChartVar.legend = new am4charts.Legend();
  nmiChartVar.legend.position = "top";
  nmiChartVar.legend.paddingBottom = 20;
  nmiChartVar.legend.labels.template.maxWidth = 95;

  var xAxis = nmiChartVar.xAxes.push(new am4charts.CategoryAxis());
  xAxis.dataFields.category = "type";
  xAxis.renderer.cellStartLocation = 0.1;
  xAxis.renderer.cellEndLocation = 0.9;
  xAxis.renderer.grid.template.location = 0;

  function getExtremes(data) {
    const initialMax = -Infinity;
    const initialMin = Infinity;

    const extremes = data.reduce(
      (acc, item) => {
        const value = parseFloat(item.total);

        if (value > acc.max) {
          acc.max = value;
        }
        if (value < acc.min) {
          acc.min = value;
        }

        return acc;
      },
      { max: initialMax, min: initialMin }
    );

    return extremes;
  }

  const { max: highestValue, min: lowestValue } = getExtremes(data);

  var yAxis = nmiChartVar.yAxes.push(new am4charts.ValueAxis());
  yAxis.min = chartId == "rmpnmibase" ? -50000 : lowestValue;
  yAxis.max = chartId == "rmpnmibase" ? 50000 : highestValue;
  yAxis.renderer.minGridDistance = 20;

  function createSeries(value, name) {
    var series = nmiChartVar.series.push(new am4charts.ColumnSeries());
    series.dataFields.valueY = value;
    series.dataFields.categoryX = "type";
    series.name = name;
    series.columns.template.tooltipText = "{name}: [bold]{valueY}[/]";
    series.columns.template.fontSize = 14;

    series.events.on("hidden", arrangeColumns);
    series.events.on("shown", arrangeColumns);

    var bullet = series.bullets.push(new am4charts.LabelBullet());
    bullet.interactionsEnabled = true;
    bullet.dy = 30;
    bullet.label.fill = am4core.color("#ffffff");

    return series;
  }

  const categories = [...new Set(data.map((item) => item.type))];
  const seriesNames = [...new Set(data.map((item) => item.margin))];

  const chartData = categories.map((category) => {
    const dataPoint = { type: category };
    seriesNames.forEach((name) => {
      const entry = data.find((d) => d.type === category && d.margin === name);
      dataPoint[name] = entry ? parseFloat(entry.total) : 0;
    });
    return dataPoint;
  });

  nmiChartVar.data = chartData;

  seriesNames.forEach((name) => {
    createSeries(name, name);
  });

  function arrangeColumns() {
    var series = nmiChartVar.series.getIndex(0);

    var w =
      1 -
      xAxis.renderer.cellStartLocation -
      (1 - xAxis.renderer.cellEndLocation);
    if (series.dataItems.length > 1) {
      var x0 = xAxis.getX(series.dataItems.getIndex(0), "categoryX");
      var x1 = xAxis.getX(series.dataItems.getIndex(1), "categoryX");
      var delta = ((x1 - x0) / nmiChartVar.series.length) * w;
      if (am4core.isNumber(delta)) {
        var middle = nmiChartVar.series.length / 2;

        var newIndex = 0;
        nmiChartVar.series.each(function (series) {
          if (!series.isHidden && !series.isHiding) {
            series.dummyData = newIndex;
            newIndex++;
          } else {
            series.dummyData = nmiChartVar.series.indexOf(series);
          }
        });
        var visibleCount = newIndex;
        var newMiddle = visibleCount / 2;

        nmiChartVar.series.each(function (series) {
          var trueIndex = nmiChartVar.series.indexOf(series);
          var newIndex = series.dummyData;

          var dx = (newIndex - trueIndex + middle - newMiddle) * delta;

          series.animate(
            { property: "dx", to: dx },
            series.interpolationDuration,
            series.interpolationEasing
          );
          series.bulletsContainer.animate(
            { property: "dx", to: dx },
            series.interpolationDuration,
            series.interpolationEasing
          );
        });
      }
    }
  }

  nmiChartVar.exporting.menu = new am4core.ExportMenu();
  nmiChartVar.events.on("validated", function () {
    loadingElement.style.display = "none";
  });
  return nmiChartVar;
};

function transformDataForChartNew(data, key) {
  return data
    .map((item) => item.data)
    .flat()
    .filter((item) => (key ? item.margin.includes(key) : true))
    .map((item) => item.type)
    .flat();
}
function transformDataForPOEChart(data, key) {
  return data
    .map((item) => item.data)
    .flat()
    .filter((item) => (key ? item.margin.includes(key) : true));
}
function totalValuesByName(data) {
  const totals = {};

  data.forEach((item) => {
    if (totals[item.name]) {
      totals[item.name] += item.value;
    } else {
      totals[item.name] = item.value;
    }
  });

  return [
    { type: "Capacity", total: totals["capacity"] || 0 },
    { type: "Commission", total: totals["commission"] || 0 },
    { type: "ESS", total: totals["ess"] || 0 },
    { type: "LGC", total: totals["lgc"] || 0 },
    { type: "Market Fees", total: totals["marketFees"] || 0 },
    { type: "Network", total: totals["network"] || 0 },
    { type: "Retail Margin", total: totals["retailMargin"] || 0 },
    { type: "Revenue", total: totals["revenue"] || 0 },
    { type: "STC", total: totals["stc"] || 0 },
    { type: "Wholesale Energy", total: totals["wholesaleEnergy"] || 0 },
  ];
}

function totalValuesByType(data) {
  const totals = [];
  const names = [
    { type: "Capacity", name: "capacity" },
    { type: "Commission", name: "commission" },
    { type: "ESS", name: "ess" },
    { type: "LGC", name: "lgc" },
    { type: "Market Fees", name: "marketFees" },
    { type: "Network", name: "network" },
    { type: "Retail Margin", name: "retailMargin" },
    { type: "Revenue", name: "revenue" },
    { type: "STC", name: "stc" },
    { type: "Wholesale Energy", name: "wholesaleEnergy" },
  ];
  data.forEach((item) => {
    item.type.forEach((typeObj) => {
      // Find if the same combination of margin and type already exists
      const existing = totals.find(
        (entry) => entry.margin === item.margin && entry.type === typeObj.name
      );

      if (existing) {
        // If it exists, add the value to the total
        existing.total += typeObj.value;
      } else {
        // If it doesn't exist, add a new entry to the array
        totals.push({
          margin: item.margin,
          type: typeObj.name,
          total: typeObj.value,
        });
      }
    });
  });
  function updateTypeValues(data) {
    return data.map((item) => {
      const foundName = names.find((n) => n.name === item.type);
      return {
        ...item,
        type: foundName ? foundName.type : item.type, // Update type if found, else keep the original
      };
    });
  }
  return updateTypeValues(totals);
}

function totalValuesByTypeNew(data) {
  const margins = {};
  const names = [
    { type: "Capacity", name: "capacity" },
    { type: "Commission", name: "commission" },
    { type: "ESS", name: "ess" },
    { type: "LGC", name: "lgc" },
    { type: "Market Fees", name: "marketFees" },
    { type: "Network", name: "network" },
    { type: "Retail Margin", name: "retailMargin" },
    { type: "Revenue", name: "revenue" },
    { type: "STC", name: "stc" },
    { type: "Wholesale Energy", name: "wholesaleEnergy" },
  ];

  // Create a mapping of type names for easy lookup
  const typeMapping = Object.fromEntries(
    names.map((item) => [item.name, item.type])
  );

  data.forEach((item) => {
    if (!margins[item.margin]) {
      // Initialize a new margin if not already present
      margins[item.margin] = {};
    }

    item.type.forEach((typeObj) => {
      // Use typeMapping to update the type value
      const updatedType = typeMapping[typeObj.name] || typeObj.name; // fallback to original name if not found

      if (margins[item.margin][updatedType]) {
        margins[item.margin][updatedType] += typeObj.value;
      } else {
        margins[item.margin][updatedType] = typeObj.value;
      }
    });
  });

  // Convert the result back to the required format
  return Object.keys(margins).map((margin) => {
    return {
      margin,
      type: Object.keys(margins[margin]).map((name) => ({
        name,
        value: margins[margin][name],
      })),
    };
  });
}

function transformData(data) {
  let result = [];

  data.forEach((item) => {
    item.data.forEach((marginData) => {
      result.push({
        margin: marginData.margin,
        type: item.nmi,
        total: marginData.value,
      });
    });
  });

  return result;
}
console.log("combinedDataJsonFull", combinedDataJsonFull);
var nmiChartVar;
var xyChartVar;
const typeList = [
  "RetailMargin",
  "Revenue",
  "Network",
  "Capacity",
  "WholesaleEnergy",
  "MarketFees",
  "ESS",
  "LGC",
  "STC",
  "Commission",
];

let filter1Data = combinedDataJsonFull;
let achievedMargin2024 = totalValuesByName(
  transformDataForChartNew(filter1Data, "FY2024")
);
let dataFy2025 = totalValuesByType(
  transformDataForPOEChart(filter1Data, "FY2025")
);
let dataFy2026 = totalValuesByType(
  transformDataForPOEChart(filter1Data, "FY2026")
);
let actualsMarginData = totalValuesByName(
  transformDataForChartNew(combinedDataJsonFull)
);

const dataRmperNMIBase = transformData(nmiJsonData);
const nmiList = combinedDataJsonFull.filter((data) => data.nmi);
var updatedDataRmperNMIBase = [];
var loadingElement = document.getElementById("loading");

// Filter 1
let updatedFilter1Data = [];
const filter1Nmi = document.getElementById("filter1-nmi");
const filter1Portfolio = document.getElementById("filter1-portfolio");
const filter1Status = document.getElementById("filter1-status");
const filter1Association = document.getElementById("filter1-association");
const filter1Agreement = document.getElementById("filter1-agreement");
setFilterData(combinedDataJsonFull, filter1Nmi, "nmi");
setFilterData(combinedDataJsonFull, filter1Portfolio, "portfolio");
setFilterData(combinedDataJsonFull, filter1Agreement, "agreement");
setFilterData(combinedDataJsonFull, filter1Status, "status");
setFilterData(combinedDataJsonFull, filter1Association, "association");

// Filter 2
const filter2Year = document.getElementById("filter2-year");
const filter2Margin = document.getElementById("filter2-margin");
getYearFilter(combinedDataJsonFull, filter2Year, "year");
getYearFilter(combinedDataJsonFull, filter2Margin, "margin");

// Filter 3
const filter3Nmi = document.getElementById("filter3-nmi");
const filter3FinancialYear = document.getElementById("filter3-fy");
const filter3Type = document.getElementById("filter3-type");
const filter3Portfolio = document.getElementById("filter3-portfolio");
const filter3Status = document.getElementById("filter3-status");
const filter3Agreement = document.getElementById("filter3-agreement");
const filter3Association = document.getElementById("filter3-association");
setFilterData(nmiJsonData, filter3Portfolio, "portfolio");
setFilterData(nmiJsonData, filter3Status, "agreement");
setFilterData(nmiJsonData, filter3Agreement, "status");
setFilterData(nmiJsonData, filter3Association, "association");
getYearFilter(nmiJsonData, filter3FinancialYear, "year");
getYearFilter(nmiJsonData, filter3Type, "nmi");

xyChart(achievedMargin2024, "hpfy24a", "Historical Performance FY2024 Actuals");
animatedPieChart(achievedMargin2024, "pfify24");
populateTable1Data(achievedMargin2024);
clusteredChart(dataFy2025, "fy25avspoe");
populateTable2Data(dataFy2025, "table2");
clusteredChart(dataFy2026, "fy26avspoe");
populateTable2Data(dataFy2026, "table3");

animatedPieChart(actualsMarginData, "acap");
populateTable3Data(actualsMarginData, "table5");

fyActualMarginVsPOE(
  totalValuesByTypeNew(transformDataForPOEChart(combinedDataJsonFull)),
  "fyamvspoe"
);
populateTable2Data(
  totalValuesByType(transformDataForPOEChart(combinedDataJsonFull)),
  "table4"
);

nmiClusteredChart(dataRmperNMIBase, "rmpnmibase");
getUniqueValue(dataRmperNMIBase, filter3Nmi);

const footerContainer = document.querySelector(".footer-container");

versionData.forEach((item) => {
  const footerItem = document.createElement("div");
  footerItem.className = "footer-item";
  footerItem.innerHTML =
    "<h3>" + item.version + "</h3>" + "<p>" + item.effectiveDate + "</p>";
  footerContainer.appendChild(footerItem);
});
function reloadChart(newData) {
  nmiClusteredChart(newData, "rmpnmibase");
  isChartLoading = true;
}

filter1Portfolio.addEventListener("change", function () {
  if (filter1Portfolio.value === "selectAll") {
    updatedFilter1Data = combinedDataJsonFull;
    for (let i = 0; i < filter1Portfolio.options.length; i++) {
      if (filter1Portfolio.options[i].value !== "selectAll") {
        filter1Portfolio.options[i].selected = true; // Select all other options
      }
    }
  } else {
    updatedFilter1Data = combinedDataJsonFull.filter(
      (data) => data.portfoio == filter1Portfolio.value
    );
  }
  filter1Data = updatedFilter1Data;
  refreshFilter1Data()
});

function refreshFilter1Data() {
  achievedMargin2024 = totalValuesByName(
    transformDataForChartNew(filter1Data, "FY2024")
  );
  dataFy2025 = totalValuesByType(
    transformDataForPOEChart(filter1Data, "FY2025")
  );
  dataFy2026 = totalValuesByType(
    transformDataForPOEChart(filter1Data, "FY2026")
  );
  xyChart(
    achievedMargin2024,
    "hpfy24a",
    "Historical Performance FY2024 Actuals"
  );
  animatedPieChart(achievedMargin2024, "pfify24");
  clusteredChart(dataFy2025, "fy25avspoe");
  populateTable2Data(dataFy2025, "table2");
  clusteredChart(dataFy2026, "fy26avspoe");
  populateTable2Data(dataFy2026, "table3");
}

filter3Nmi.addEventListener("change", function () {
  if (filter3Nmi.value === "selectAll") {
    updatedDataRmperNMIBase = dataRmperNMIBase;
    for (let i = 0; i < filter3Nmi.options.length; i++) {
      if (filter3Nmi.options[i].value !== "selectAll") {
        filter3Nmi.options[i].selected = true; // Select all other options
      }
    }
  } else {
    updatedDataRmperNMIBase = transformData(
      nmiJsonData.filter(
        (data) =>
          data.typesofmargin != "(blank)" && data.typesofmargin != "Grand Total"
      )
    ).filter((data) => data.type == filter3Nmi.value);
  }
  reloadChart(updatedDataRmperNMIBase);
  console.log("dataRmperNMIBase", updatedDataRmperNMIBase, filter3Nmi.value);
});

function toggleSidebar() {
  var sidebar = document.getElementById("mySidebar");
  var width = sidebar.style.width;

  if (width === "250px") {
    sidebar.style.width = "0";
  } else {
    sidebar.style.width = "250px";
  }
}
