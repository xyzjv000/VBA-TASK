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

  return chart;
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

  return chart;
};

const populateTable1Data = (data) => {
  const headerRow = document.querySelector("#table1 #headerRow");
  const dataRow = document.querySelector("#table1 #dataRow");

  // Clear the existing header and data rows
  headerRow.innerHTML = "";
  dataRow.innerHTML = "";

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
  return chart;
};

const populateTable2Data = (data, id) => {
  const headerRow = document.querySelector("#" + id + " #headerRow");
  const tbody = document.querySelector("#" + id + " tbody");

  // Clear the existing header and body content
  headerRow.innerHTML = "";
  tbody.innerHTML = "";

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

  return chart;
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
  if (chartId == "rmpnmibase") {
    updateHeaderData(data);
    console.log(nmiBaseChartData);
  }
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

function reloadChart(newData) {
  nmiClusteredChart(newData, "rmpnmibase");
  isChartLoading = true;
}

function refreshFilter1Data(filteredData) {
  achievedMargin2024 = totalValuesByName(
    transformDataForChartNew(filteredData, "FY2024")
  );
  dataFy2025 = totalValuesByType(
    transformDataForPOEChart(filteredData, "FY2025")
  );
  dataFy2026 = totalValuesByType(
    transformDataForPOEChart(filteredData, "FY2026")
  );
  if (xyChartHpfy24a) {
    xyChartHpfy24a.dispose();
  }
  if (clusterChartFy25avspoe) {
    clusterChartFy25avspoe.dispose();
  }

  if (clusterChartFy26avspoe) {
    clusterChartFy26avspoe.dispose();
  }

  xyChartHpfy24a = xyChart(
    achievedMargin2024,
    "hpfy24a",
    "Historical Performance FY2024 Actuals"
  );
  // animatedPieChart(achievedMargin2024, "pfify24");
  populateTable1Data(achievedMargin2024);
  clusterChartFy25avspoe = clusteredChart(dataFy2025, "fy25avspoe");
  populateTable2Data(dataFy2025, "table2");
  clusterChartFy26avspoe = clusteredChart(dataFy2026, "fy26avspoe");
  populateTable2Data(dataFy2026, "table3");
}

// Function to apply all active filters
// Function to apply all active filters
function applyAllFilters() {
  // Start with the full dataset before filtering
  let filteredData = combinedDataJsonFull;

  // Apply portfolio filter if it's not 'selectAll'
  if (filter1Portfolio.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.portfolio === filter1Portfolio.value
    );
  }

  // Apply status filter if it's not 'selectAll'
  if (filter1Status.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.status === filter1Status.value
    );
  }

  // Apply association filter if it's not 'selectAll'
  if (filter1Association.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.association === filter1Association.value
    );
  }

  // Apply agreement filter if it's not 'selectAll'
  if (filter1Agreement.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.agreement === filter1Agreement.value
    );
  }

  // Apply NMI filter if it's not 'selectAll'
  if (filter1Nmi.value !== "selectAll") {
    filteredData = filteredData.filter((data) => data.nmi === filter1Nmi.value);
  }

  // Ensure filtered data is updated globally and reflect the final filtered result
  filter1Data = filteredData;

  // setFilterData(filteredData, filter1Nmi, "nmi");
  // setFilterData(filteredData, filter1Portfolio, "portfolio");
  // setFilterData(filteredData, filter1Agreement, "agreement");
  // setFilterData(filteredData, filter1Status, "status");
  // setFilterData(filteredData, filter1Association, "association");

  // Refresh table or chart with the final filtered data
  refreshFilter1Data(filteredData);
}

function refreshFilter2Data(filteredData) {
  actualsMarginData = totalValuesByName(transformDataForChartNew(filteredData));

  if (animatedPieChartAcap) {
    animatedPieChartAcap.dispose();
  }

  if (fyAMVP) {
    fyAMVP.dispose();
  }

  animatedPieChartAcap = animatedPieChart(actualsMarginData, "acap");
  populateTable3Data(actualsMarginData, "table5");

  fyAMVP = fyActualMarginVsPOE(
    totalValuesByTypeNew(transformDataForPOEChart(filteredData)),
    "fyamvspoe"
  );
  populateTable2Data(
    totalValuesByType(transformDataForPOEChart(filteredData)),
    "table4"
  );
}

function applyAllFilters2() {
  // Start with the full dataset before filtering
  let filteredData = combinedDataJsonFull;

  // Apply portfolio filter if it's not 'selectAll'
  if (filter2Year.value !== "selectAll") {
    filteredData = filteredData.map((data) => {
      return {
        ...data,
        data: data.data.filter((item) =>
          item.margin.includes(filter2Year.value)
        ),
      };
    });
  }

  if (filter2Margin.value !== "selectAll") {
    filteredData = filteredData.map((data) => {
      return {
        ...data,
        data: data.data.filter((item) => item.margin == filter2Margin.value),
      };
    });
  }

  // Ensure filtered data is updated globally and reflect the final filtered result
  filter2Data = filteredData;
  console.log("Filtered data:", filter2Data);
  // getYearFilter(filteredData, filter2Year, "year");
  // getYearFilter(filteredData, filter2Margin, "margin");
  refreshFilter2Data(filteredData);
}

function refreshFilter3Data(filteredData) {
  reloadChart(transformData(filteredData));
}

function applyAllFilters3() {
  // Start with the full dataset before filtering
  let filteredData = nmiJsonData;
  // Apply NMI filter if it's not 'selectAll'
  if (filter3Nmi.value !== "selectAll") {
    filteredData = filteredData.filter((data) => data.nmi === filter3Nmi.value);
  }

  if (filter3FinancialYear.value !== "selectAll") {
    filteredData = filteredData.map((data) => {
      return {
        ...data,
        data: data.data.filter((item) =>
          item.margin.includes(filter3FinancialYear.value)
        ),
      };
    });
  }

  if (filter3Type.value !== "selectAll") {
    filteredData = filteredData.map((data) => {
      return {
        ...data,
        data: data.data.filter((item) => item.margin == filter3Type.value),
      };
    });
  }

  if (filter3Portfolio.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.portfolio === filter3Portfolio.value
    );
  }
  if (filter3Status.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.status === filter3Status.value
    );
  }
  if (filter3Agreement.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.agreement === filter3Agreement.value
    );
  }
  if (filter3Association.value !== "selectAll") {
    filteredData = filteredData.filter(
      (data) => data.association === filter3Association.value
    );
  }

  // Ensure filtered data is updated globally and reflect the final filtered result
  filter3Data = filteredData;
  refreshFilter3Data(filteredData);
}

function toggleSidebar() {
  var sidebar = document.getElementById("mySidebar");
  var width = sidebar.style.width;

  if (width === "250px") {
    sidebar.style.width = "0";
  } else {
    sidebar.style.width = "250px";
  }
}
let currentPage = "page1";
function updateSidebarFilters() {
  // Get sections
  var section1 = document.getElementById("page1");
  var section2 = document.getElementById("page2");

  // Get filters
  var filter0Height = document.getElementById("filter0").offsetHeight;
  var filter1 = document.getElementById("filter1");
  var filter2 = document.getElementById("filter2");
  var filter3 = document.getElementById("filter3");
  var scrollPosition = (window.scrollY || window.pageYOffset) + filter0Height;

  if (scrollPosition <= section1.offsetHeight) {
    filter1.style.display = "flex";
    filter2.style.display = "none";
    filter3.style.display = "none";
    currentPage = "page1";
  } else if (
    scrollPosition > section1.offsetHeight &&
    scrollPosition <= section1.offsetHeight + section2.offsetHeight
  ) {
    filter1.style.display = "none";
    filter2.style.display = "flex";
    filter3.style.display = "none";
    currentPage = "page2";
  } else {
    filter1.style.display = "none";
    filter2.style.display = "none";
    filter3.style.display = "flex";
    currentPage = "page3";
  }
}
function updateHeaderData(data) {
  nmiBaseChartData = {
    achievedMarginFy2024: 0,
    actualMarginFy2025: 0,
    actualMarginFy2026: 0,
    predictedMarginFy2025: 0,
    predictedMarginFy2026: 0,
  };

  nmiBaseChartData.achievedMarginFy2024 = data
    .filter((item) => item.margin == "Achieved Margin FY2024")
    .reduce((sum, item) => sum + item.total, 0);
  nmiBaseChartData.actualMarginFy2025 = data
    .filter((item) => item.margin == "Total Actuals Margin FY2025")
    .reduce((sum, item) => sum + item.total, 0);
  nmiBaseChartData.actualMarginFy2026 = data
    .filter((item) => item.margin == "Total Actuals Margin FY2026")
    .reduce((sum, item) => sum + item.total, 0);
  nmiBaseChartData.predictedMarginFy2025 = data
    .filter((item) => item.margin == "Total Predicted TM FY2025")
    .reduce((sum, item) => sum + item.total, 0);
  nmiBaseChartData.predictedMarginFy2026 = data
    .filter((item) => item.margin == "Total Predicted TM FY2026")
    .reduce((sum, item) => sum + item.total, 0);

  let items = document.querySelectorAll(".header-items");
  items.forEach((item) => {
    const label = item.querySelector(".header-item-label").textContent.trim();
    const valueElement = item.querySelector(".header-item-value");

    if (label === "Achieved Margin FY2024") {
      valueElement.textContent = `$${nmiBaseChartData.achievedMarginFy2024.toLocaleString()}`;
    } else if (label === "Actual Margin FY2025") {
      valueElement.textContent = `$${nmiBaseChartData.actualMarginFy2025.toLocaleString()}`;
    } else if (label === "Actual Margin FY2026") {
      valueElement.textContent = `$${nmiBaseChartData.actualMarginFy2026.toLocaleString()}`;
    } else if (label === "Predicted Margin FY 2025") {
      valueElement.textContent = `$${nmiBaseChartData.predictedMarginFy2025.toLocaleString()}`;
    } else if (label === "Predicted Margin FY 2026") {
      valueElement.textContent = `$${nmiBaseChartData.predictedMarginFy2026.toLocaleString()}`;
    }
  });
}
updateSidebarFilters();

let nmiBaseChartData = {
  achievedMarginFy2024: 0,
  actualMarginFy2025: 0,
  actualMarginFy2026: 0,
  predictedMarginFy2025: 0,
  predictedMarginFy2026: 0,
};
console.log("combinedDataJsonFull", combinedDataJsonFull);
var nmiChartVar;
var xyChartVar;

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
let filter3Data = nmiJsonData;
const dataRmperNMIBase = transformData(filter3Data);
const nmiList = combinedDataJsonFull.filter((data) => data.nmi);
var updatedDataRmperNMIBase = [];
var loadingElement = document.getElementById("loading");

// Filter 1
let updatedFilter1Data = [];
let updatedFilter2Data = [];

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
setFilterData(nmiJsonData, filter3Status, "status");
setFilterData(nmiJsonData, filter3Agreement, "agreement");
setFilterData(nmiJsonData, filter3Association, "association");
getYearFilter(nmiJsonData, filter3FinancialYear, "year");
getYearFilter(nmiJsonData, filter3Type, "margin");

const clearFilterButton = document.querySelector("#clearFilters button");
const exportButton = document.querySelector("#exportButton button");
const exportDataButton = document.querySelector("#exportDataButton button");

let xyChartHpfy24a = xyChart(
  achievedMargin2024,
  "hpfy24a",
  "Historical Performance FY2024 Actuals"
);
animatedPieChart(achievedMargin2024, "pfify24");
populateTable1Data(achievedMargin2024);
let clusterChartFy25avspoe = clusteredChart(dataFy2025, "fy25avspoe");
populateTable2Data(dataFy2025, "table2");
let clusterChartFy26avspoe = clusteredChart(dataFy2026, "fy26avspoe");
populateTable2Data(dataFy2026, "table3");

let filter2Data = combinedDataJsonFull;
let actualsMarginData = totalValuesByName(
  transformDataForChartNew(filter2Data)
);
let animatedPieChartAcap = animatedPieChart(actualsMarginData, "acap");
populateTable3Data(actualsMarginData, "table5");

let fyAMVP = fyActualMarginVsPOE(
  totalValuesByTypeNew(transformDataForPOEChart(filter2Data)),
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
// event listeners
// Handle the change events for Filter 1
filter1Portfolio.addEventListener("change", function () {
  if (filter1Portfolio.value === "selectAll") {
    // If "selectAll" is chosen, reset to full dataset, but apply other filters
    filter1Portfolio.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters();
});

filter1Status.addEventListener("change", function () {
  if (filter1Status.value === "selectAll") {
    filter1Status.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters();
});

filter1Association.addEventListener("change", function () {
  if (filter1Association.value === "selectAll") {
    filter1Association.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters();
});

filter1Agreement.addEventListener("change", function () {
  if (filter1Agreement.value === "selectAll") {
    filter1Agreement.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters();
});

filter1Nmi.addEventListener("change", function () {
  if (filter1Nmi.value === "selectAll") {
    filter1Nmi.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters();
});

// Handle the change events for Filter 2
filter2Year.addEventListener("change", function () {
  if (filter2Year.value === "selectAll") {
    filter2Year.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters2();
});

filter2Margin.addEventListener("change", function () {
  if (filter2Margin.value === "selectAll") {
    filter2Margin.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters2();
});

filter3Nmi.addEventListener("change", function () {
  if (filter3Nmi.value === "selectAll") {
    filter3Nmi.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

filter3FinancialYear.addEventListener("change", function () {
  if (filter3FinancialYear.value === "selectAll") {
    filter3FinancialYear.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

filter3Type.addEventListener("change", function () {
  if (filter3Type.value === "selectAll") {
    filter3Type.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

filter3Portfolio.addEventListener("change", function () {
  if (filter3Portfolio.value === "selectAll") {
    filter3Portfolio.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

filter3Status.addEventListener("change", function () {
  if (filter3Status.value === "selectAll") {
    filter3Status.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

filter3Agreement.addEventListener("change", function () {
  if (filter3Agreement.value === "selectAll") {
    filter3Agreement.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

filter3Association.addEventListener("change", function () {
  if (filter3Association.value === "selectAll") {
    filter3Association.value = "selectAll"; // Reset this filter to 'selectAll'
  }
  applyAllFilters3();
});

window.addEventListener("scroll", updateSidebarFilters);

clearFilterButton.addEventListener("click", () => {
  // Check if any filter is NOT 'selectAll'
  const filter1Changed =
    filter1Nmi.value !== "selectAll" ||
    filter1Portfolio.value !== "selectAll" ||
    filter1Status.value !== "selectAll" ||
    filter1Association.value !== "selectAll" ||
    filter1Agreement.value !== "selectAll";

  const filter2Changed =
    filter2Year.value !== "selectAll" || filter2Margin.value !== "selectAll";

  const filter3Changed =
    filter3Nmi.value !== "selectAll" ||
    filter3FinancialYear.value !== "selectAll" ||
    filter3Type.value !== "selectAll" ||
    filter3Portfolio.value !== "selectAll" ||
    filter3Status.value !== "selectAll" ||
    filter3Agreement.value !== "selectAll" ||
    filter3Association.value !== "selectAll";
  // Set all filters to 'selectAll'
  filter1Nmi.value = "selectAll";
  filter1Portfolio.value = "selectAll";
  filter1Status.value = "selectAll";
  filter1Association.value = "selectAll";
  filter1Agreement.value = "selectAll";

  filter2Year.value = "selectAll";
  filter2Margin.value = "selectAll";

  filter3Nmi.value = "selectAll";
  filter3FinancialYear.value = "selectAll";
  filter3Type.value = "selectAll";
  filter3Portfolio.value = "selectAll";
  filter3Status.value = "selectAll";
  filter3Agreement.value = "selectAll";
  filter3Association.value = "selectAll";

  console.log("filter1Changed:", filter1Changed);
  console.log("filter2Changed:", filter2Changed);
  console.log("filter3Changed:", filter3Changed);

  // Only apply the filters if any of the corresponding filter sets are changed
  if (filter1Changed) {
    applyAllFilters();
  }
  if (filter2Changed) {
    applyAllFilters2();
  }
  if (filter3Changed) {
    applyAllFilters3();
  }
});

exportButton.addEventListener("click", () => {
  const captureElement = document.getElementById(currentPage);

  // Use dom-to-image to capture the element as a PNG image
  domtoimage
    .toPng(captureElement)
    .then((dataUrl) => {
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF();

      const img = new Image();
      img.src = dataUrl;

      img.onload = function () {
        const imgWidth = 210; // A4 size width in mm (210mm)
        const imgHeight = (img.height * imgWidth) / img.width; // Maintain aspect ratio

        // Add image to PDF
        pdf.addImage(img, "PNG", 0, 0, imgWidth, imgHeight);

        // Save the PDF
        pdf.save("ExportedReportCharts.pdf");
      };
    })
    .catch((error) => {
      console.error("Error capturing element:", error);
    });
});

exportDataButton.addEventListener("click", () => {
  const chartIds = ["hpfy24a", "pfify24", "fy25avspoe", "fy26avspoe"];
});

// // Add event listeners if needed for actions when buttons are selected
// document.getElementById("mtd").addEventListener("change", function () {
//   if (this.checked) {
//     console.log("MTD selected");
//   }
// });

// document.getElementById("ytd").addEventListener("change", function () {
//   if (this.checked) {
//     console.log("YTD selected");
//   }
// });
