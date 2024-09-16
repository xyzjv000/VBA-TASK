Dim vbaCode As String
vbaCode = "" & vbCrLf & _
"<!DOCTYPE html>" & vbCrLf & _
"<html lang=""en"">" & vbCrLf & _
"  <head>" & vbCrLf & _
"    <meta charset=""UTF-8"" />" & vbCrLf & _
"    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />" & vbCrLf & _
"    <title>Chart Report</title>" & vbCrLf & _
"    <style>" & vbCrLf & _
"      body {" & vbCrLf & _
"        font-family: Arial, sans-serif;" & vbCrLf & _
"        margin: 0;" & vbCrLf & _
"        padding: 0;" & vbCrLf & _
"        display: flex;" & vbCrLf & _
"        flex-direction: column;" & vbCrLf & _
"        justify-content: center;" & vbCrLf & _
"        align-items: center;" & vbCrLf & _
"        background-color: #f4f4f4;" & vbCrLf & _
"      }" & vbCrLf & _
"      .container {" & vbCrLf & _
"        width: 90%;" & vbCrLf & _
"        margin: auto;" & vbCrLf & _
"        padding: 20px;" & vbCrLf & _
"        background: #fff;" & vbCrLf & _
"        border-radius: 8px;" & vbCrLf & _
"        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);" & vbCrLf & _
"      }" & vbCrLf & _
"      img {" & vbCrLf & _
"        max-width: 100%;" & vbCrLf & _
"        height: auto;" & vbCrLf & _
"      }" & vbCrLf & _
"      table {" & vbCrLf & _
"        width: 100%;" & vbCrLf & _
"        border-collapse: collapse;" & vbCrLf & _
"        margin-top: 20px;" & vbCrLf & _
"      }" & vbCrLf & _
"      table," & vbCrLf & _
"      th," & vbCrLf & _
"      td {" & vbCrLf & _
"        border: 1px solid #ddd;" & vbCrLf & _
"      }" & vbCrLf & _
"      th," & vbCrLf & _
"      td {" & vbCrLf & _
"        padding: 10px;" & vbCrLf & _
"        text-align: left;" & vbCrLf & _
"      }" & vbCrLf & _
"      th {" & vbCrLf & _
"        background-color: #f2f2f2;" & vbCrLf & _
"      }" & vbCrLf & _
"      #hpfy24a," & vbCrLf & _
"      #pfify24," & vbCrLf & _
"      #fy25avspoe," & vbCrLf & _
"      #fy26avspoe {" & vbCrLf & _
"        height: 400px;" & vbCrLf & _
"      }" & vbCrLf & _
"      #acap {" & vbCrLf & _
"        height: 500px;" & vbCrLf & _
"        width: 70%;" & vbCrLf & _
"        min-width: 600px;" & vbCrLf & _
"      }" & vbCrLf & _
"" & vbCrLf & _
"      #fyamvspoe {" & vbCrLf & _
"        height: 600px;" & vbCrLf & _
"      }" & vbCrLf & _
"" & vbCrLf & _
"      .chart-container {" & vbCrLf & _
"        display: flex;" & vbCrLf & _
"        flex-wrap: wrap;" & vbCrLf & _
"      }" & vbCrLf & _
"" & vbCrLf & _
"      .chart-item {" & vbCrLf & _
"        flex: 1;" & vbCrLf & _
"        min-width: 400px;" & vbCrLf & _
"      }" & vbCrLf & _
"" & vbCrLf & _
"      #rmpnmibase {" & vbCrLf & _
"        min-height: 700px;" & vbCrLf & _
"      }" & vbCrLf & _
"      #nmi {" & vbCrLf & _
"        width: 200px;" & vbCrLf & _
"        padding: 10px;" & vbCrLf & _
"        border-radius: 5px;" & vbCrLf & _
"      }" & vbCrLf & _
"      #loading {" & vbCrLf & _
"        text-align: center;" & vbCrLf & _
"      }" & vbCrLf & _
"    </style>" & vbCrLf & _
"    <style>" & vbCrLf & _
"      body {" & vbCrLf & _
"        font-family: Arial, sans-serif;" & vbCrLf & _
"        margin: 0;" & vbCrLf & _
"        padding: 0;" & vbCrLf & _
"      }" & vbCrLf & _
"      .footer {" & vbCrLf & _
"        background-color: #f1f1f1;" & vbCrLf & _
"        padding: 20px;" & vbCrLf & _
"        border-top: 1px solid #ddd;" & vbCrLf & _
"      }" & vbCrLf & _
"      .footer-container {" & vbCrLf & _
"        margin: 0 auto;" & vbCrLf & _
"        display: flex;" & vbCrLf & _
"        flex-wrap: wrap;" & vbCrLf & _
"        gap: 20px;" & vbCrLf & _
"      }" & vbCrLf & _
"      .footer-item {" & vbCrLf & _
"        background-color: #ffffff;" & vbCrLf & _
"        border: 1px solid #ddd;" & vbCrLf & _
"        border-radius: 8px;" & vbCrLf & _
"        padding: 15px;" & vbCrLf & _
"        flex: 1;" & vbCrLf & _
"        min-width: 200px;" & vbCrLf & _
"      }" & vbCrLf & _
"      .footer-item h3 {" & vbCrLf & _
"        margin: 0 0 10px;" & vbCrLf & _
"        font-size: 16px;" & vbCrLf & _
"        color: #333;" & vbCrLf & _
"      }" & vbCrLf & _
"      .footer-item p {" & vbCrLf & _
"        margin: 0;" & vbCrLf & _
"        font-size: 14px;" & vbCrLf & _
"        color: #666;" & vbCrLf & _
"      }" & vbCrLf & _
"      @media (max-width: 768px) {" & vbCrLf & _
"        .footer-item {" & vbCrLf & _
"          flex: 1 1 100%;" & vbCrLf & _
"        }" & vbCrLf & _
"      }" & vbCrLf & _
"    </style>" & vbCrLf & _
"    <style>" & vbCrLf & _
"      /* Responsive styles */" & vbCrLf & _
"      @media (max-width: 768px) {" & vbCrLf & _
"        #table1 thead, #table1 tbody {" & vbCrLf & _
"          display: inline-grid;" & vbCrLf & _
"          width: 50%;" & vbCrLf & _
"        }" & vbCrLf & _
"        #table1 #headerRow, #table1 #dataRow {" & vbCrLf & _
"          display: inline-grid;" & vbCrLf & _
"        }" & vbCrLf & _
"        .table-container {" & vbCrLf & _
"          overflow: scroll;" & vbCrLf & _
"        }" & vbCrLf & _
"        " & vbCrLf & _
"        .container{" & vbCrLf & _
"          width: 98%;" & vbCrLf & _
"        }" & vbCrLf & _
"      }" & vbCrLf & _
"    </style>" & vbCrLf & _
"  </head>" & vbCrLf & _
"  <body>" & vbCrLf & _
"    <div class=""container"">" & vbCrLf & _
"      <div class=""chart"">" & vbCrLf & _
"        <div class=""chart-container"">" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>Historical Performance FY2024 Actuals</h3></center>" & vbCrLf & _
"            <div id=""hpfy24a""></div>" & vbCrLf & _
"          </div>" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>Percentage Fees in FY2024</h3></center>" & vbCrLf & _
"            <div id=""pfify24""></div>" & vbCrLf & _
"          </div>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <div class=""table-container"">" & vbCrLf & _
"          <table id=""table1"">" & vbCrLf & _
"            <thead>" & vbCrLf & _
"              <tr id=""headerRow""></tr>" & vbCrLf & _
"            </thead>" & vbCrLf & _
"            <tbody>" & vbCrLf & _
"              <tr id=""dataRow""></tr>" & vbCrLf & _
"            </tbody>" & vbCrLf & _
"          </table>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <br />" & vbCrLf & _
"        <hr />" & vbCrLf & _
"        <!-- FY25 Actual vs POE -->" & vbCrLf & _
"        <div>" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>FY25 Actual vs POE</h3></center>" & vbCrLf & _
"            <div id=""fy25avspoe""></div>" & vbCrLf & _
"          </div>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <div class=""table-container"">" & vbCrLf & _
"          <table id=""table2"">" & vbCrLf & _
"            <thead>" & vbCrLf & _
"              <tr id=""headerRow""></tr>" & vbCrLf & _
"            </thead>" & vbCrLf & _
"            <tbody>" & vbCrLf & _
"              <tr id=""dataRow""></tr>" & vbCrLf & _
"            </tbody>" & vbCrLf & _
"          </table>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <br />" & vbCrLf & _
"        <hr />" & vbCrLf & _
"        <!-- FY26 Actual vs POE -->" & vbCrLf & _
"        <div>" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>FY26 Actual vs POE</h3></center>" & vbCrLf & _
"            <div id=""fy26avspoe""></div>" & vbCrLf & _
"          </div>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <div class=""table-container"">" & vbCrLf & _
"          <table id=""table3"">" & vbCrLf & _
"            <thead>" & vbCrLf & _
"              <tr id=""headerRow""></tr>" & vbCrLf & _
"            </thead>" & vbCrLf & _
"            <tbody>" & vbCrLf & _
"              <tr id=""dataRow""></tr>" & vbCrLf & _
"            </tbody>" & vbCrLf & _
"          </table>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <br />" & vbCrLf & _
"        <hr />" & vbCrLf & _
"        <!-- Actual Cost and Percentage -->" & vbCrLf & _
"        <div>" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>Actual Cost and Percentage</h3></center>" & vbCrLf & _
"            <div style=""display: flex; flex-wrap: wrap; gap: 1rem"">" & vbCrLf & _
"              <div id=""acap""></div>" & vbCrLf & _
"              <div style=""flex: 1"" class=""table-container"">" & vbCrLf & _
"                <table id=""table5"">" & vbCrLf & _
"                  <thead>" & vbCrLf & _
"                    <tr id=""headerRow""></tr>" & vbCrLf & _
"                  </thead>" & vbCrLf & _
"                  <tbody>" & vbCrLf & _
"                    <tr id=""dataRow""></tr>" & vbCrLf & _
"                  </tbody>" & vbCrLf & _
"                </table>" & vbCrLf & _
"              </div>" & vbCrLf & _
"            </div>" & vbCrLf & _
"          </div>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <br />" & vbCrLf & _
"        <hr />" & vbCrLf & _
"        <div>" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>Financial Year Actual Margin vs POE</h3></center>" & vbCrLf & _
"            <div id=""fyamvspoe""></div>" & vbCrLf & _
"          </div>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <div class=""table-container"">" & vbCrLf & _
"          <table id=""table4"">" & vbCrLf & _
"            <thead>" & vbCrLf & _
"              <tr id=""headerRow""></tr>" & vbCrLf & _
"            </thead>" & vbCrLf & _
"            <tbody>" & vbCrLf & _
"              <tr id=""dataRow""></tr>" & vbCrLf & _
"            </tbody>" & vbCrLf & _
"          </table>" & vbCrLf & _
"        </div>" & vbCrLf & _
"        <hr />" & vbCrLf & _
"        <br />" & vbCrLf & _
"        <!-- Retail Margin per NMI Base -->" & vbCrLf & _
"        <div>" & vbCrLf & _
"          <div class=""chart-item"">" & vbCrLf & _
"            <center><h3>Retail Margin per NMI Base</h3></center>" & vbCrLf & _
"            <h4 id=""loading"">Loading please wait...</h4>" & vbCrLf & _
"            <div id=""rmpnmibase""></div>" & vbCrLf & _
"            <label for=""nmi"">Filter by NMI</label>" & vbCrLf & _
"            <select name=""nmi"" id=""nmi""></select>" & vbCrLf & _
"          </div>" & vbCrLf & _
"        </div>" & vbCrLf & _
"      </div>" & vbCrLf & _
"    </div>" & vbCrLf & _
"    <div class=""footer"">" & vbCrLf & _
"      <div class=""footer-container"">" & vbCrLf & _
"        <!-- Footer items will be injected here -->" & vbCrLf & _
"      </div>" & vbCrLf & _
"    </div>" & vbCrLf & _
"    <script src=""https://cdn.amcharts.com/lib/4/core.js""></script>" & vbCrLf & _
"    <script src=""https://cdn.amcharts.com/lib/4/charts.js""></script>" & vbCrLf & _
"    <script src=""https://cdn.amcharts.com/lib/4/maps.js""></script>" & vbCrLf & _
"    <script src=""https://cdn.amcharts.com/lib/4/themes/animated.js""></script>" & vbCrLf & _
"    <!-- functions -->" & vbCrLf & _
"    <script>" & vbCrLf & _
"      const transformDataForChart = (data, marginType) => {" & vbCrLf & _
"        const marginData = data.find((entry) =>" & vbCrLf & _
"          entry.typesofmargin.includes(marginType)" & vbCrLf & _
"        );" & vbCrLf & _
"        if (!marginData) {" & vbCrLf & _
"          console.error(""No data found for the specified typesofmargin."");" & vbCrLf & _
"          return [];" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        return Object.keys(marginData.totals)" & vbCrLf & _
"          .filter((key) => key.toLowerCase() !== ""total"")" & vbCrLf & _
"          .map((key) => ({" & vbCrLf & _
"            type: marginData.labels[key] || key," & vbCrLf & _
"            total: parseFloat(marginData.totals[key])," & vbCrLf & _
"          }));" & vbCrLf & _
"      };" & vbCrLf & _
"      const transformDataForCluster = (data, marginType) => {" & vbCrLf & _
"        var marginData = [];" & vbCrLf & _
"        if (marginType == null || marginType == """") {" & vbCrLf & _
"          marginData = data;" & vbCrLf & _
"        } else {" & vbCrLf & _
"          marginData = data.filter((entry) =>" & vbCrLf & _
"            entry.typesofmargin.includes(marginType)" & vbCrLf & _
"          );" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        const combineData = (data) => {" & vbCrLf & _
"          const combinedArray = [];" & vbCrLf & _
"" & vbCrLf & _
"          data.forEach((entry) => {" & vbCrLf & _
"            const { typesofmargin, totals, labels } = entry;" & vbCrLf & _
"" & vbCrLf & _
"            Object.keys(totals).forEach((key) => {" & vbCrLf & _
"              if (key.toLowerCase() !== ""total"") {" & vbCrLf & _
"                combinedArray.push({" & vbCrLf & _
"                  margin: typesofmargin," & vbCrLf & _
"                  type: labels[key] || key," & vbCrLf & _
"                  total: parseFloat(totals[key])," & vbCrLf & _
"                });" & vbCrLf & _
"              }" & vbCrLf & _
"            });" & vbCrLf & _
"          });" & vbCrLf & _
"" & vbCrLf & _
"          return combinedArray;" & vbCrLf & _
"        };" & vbCrLf & _
"" & vbCrLf & _
"        if (!marginData) {" & vbCrLf & _
"          console.error(""No data found for the specified typesofmargin."");" & vbCrLf & _
"          return [];" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        return combineData(marginData);" & vbCrLf & _
"      };" & vbCrLf & _
"      const animatedPieChart = (data, chartId) => {" & vbCrLf & _
"        am4core.useTheme(am4themes_animated);" & vbCrLf & _
"        var chart = am4core.create(chartId, am4charts.PieChart3D);" & vbCrLf & _
"" & vbCrLf & _
"        chart.data = data;" & vbCrLf & _
"        chart.innerRadius = am4core.percent(40);" & vbCrLf & _
"" & vbCrLf & _
"        var pieSeries = chart.series.push(new am4charts.PieSeries3D());" & vbCrLf & _
"        pieSeries.dataFields.value = ""total"";" & vbCrLf & _
"        pieSeries.dataFields.category = ""type"";" & vbCrLf & _
"" & vbCrLf & _
"        pieSeries.labels.template.disabled = chartId != ""acap"";" & vbCrLf & _
"        pieSeries.ticks.template.disabled = chartId != ""acap"";" & vbCrLf & _
"" & vbCrLf & _
"        chart.legend = new am4charts.Legend();" & vbCrLf & _
"        chart.legend.position = ""right"";" & vbCrLf & _
"        chart.exporting.menu = new am4core.ExportMenu();" & vbCrLf & _
"      };" & vbCrLf & _
"      const xyChart = (data, chartId, seriesName) => {" & vbCrLf & _
"        am4core.useTheme(am4themes_animated);" & vbCrLf & _
"        var chart = am4core.create(chartId, am4charts.XYChart);" & vbCrLf & _
"" & vbCrLf & _
"        chart.data = data;" & vbCrLf & _
"        var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());" & vbCrLf & _
"        categoryAxis.dataFields.category = ""type"";" & vbCrLf & _
"        categoryAxis.title.text = ""Margins"";" & vbCrLf & _
"" & vbCrLf & _
"        categoryAxis.renderer.minGridDistance = 20;" & vbCrLf & _
"" & vbCrLf & _
"        categoryAxis.renderer.labels.template.adapter.add(" & vbCrLf & _
"          ""dy""," & vbCrLf & _
"          function (dy, target) {" & vbCrLf & _
"            return 0;" & vbCrLf & _
"          }" & vbCrLf & _
"        );" & vbCrLf & _
"" & vbCrLf & _
"        categoryAxis.renderer.labels.template.rotation = 45;" & vbCrLf & _
"        categoryAxis.renderer.labels.template.horizontalCenter = ""right"";" & vbCrLf & _
"        categoryAxis.renderer.labels.template.verticalCenter = ""middle"";" & vbCrLf & _
"        categoryAxis.renderer.labels.template.fontSize = 12;" & vbCrLf & _
"" & vbCrLf & _
"        var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());" & vbCrLf & _
"        valueAxis.title.text = ""Value"";" & vbCrLf & _
"" & vbCrLf & _
"        var series = chart.series.push(new am4charts.ColumnSeries());" & vbCrLf & _
"        series.dataFields.valueY = ""total"";" & vbCrLf & _
"        series.dataFields.categoryX = ""type"";" & vbCrLf & _
"        series.name = seriesName;" & vbCrLf & _
"        series.stacked = true;" & vbCrLf & _
"" & vbCrLf & _
"        series.columns.template.tooltipText = ""{categoryX}: [bold]{valueY}[/]"";" & vbCrLf & _
"        series.columns.template.fontSize = 14;" & vbCrLf & _
"" & vbCrLf & _
"        series.tooltip.fontSize = 12;" & vbCrLf & _
"" & vbCrLf & _
"        chart.yAxes.getIndex(0).renderer.labels.template.fontSize = 12;" & vbCrLf & _
"        chart.xAxes.getIndex(0).renderer.labels.template.fontSize = 12;" & vbCrLf & _
"" & vbCrLf & _
"        chart.cursor = new am4charts.XYCursor();" & vbCrLf & _
"        chart.exporting.menu = new am4core.ExportMenu();" & vbCrLf & _
"      };" & vbCrLf & _
"      const populateTable1Data = (data) => {" & vbCrLf & _
"        const headerRow = document.querySelector(""#table1 #headerRow"");" & vbCrLf & _
"        const dataRow = document.querySelector(""#table1 #dataRow"");" & vbCrLf & _
"" & vbCrLf & _
"        data.forEach((item) => {" & vbCrLf & _
"          const th = document.createElement(""th"");" & vbCrLf & _
"          th.textContent = item.type" & vbCrLf & _
"            .replace(/([A-Z][a-z]+|[A-Z]+(?![a-z]))/g, "" $1"")" & vbCrLf & _
"            .trim();" & vbCrLf & _
"          headerRow.appendChild(th);" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        data.forEach((item) => {" & vbCrLf & _
"          const td = document.createElement(""td"");" & vbCrLf & _
"          td.textContent = parseFloat(item.total).toFixed(2);" & vbCrLf & _
"          dataRow.appendChild(td);" & vbCrLf & _
"        });" & vbCrLf & _
"      };" & vbCrLf & _
"      const clusteredChart = (data, chartId) => {" & vbCrLf & _
"        am4core.useTheme(am4themes_animated);" & vbCrLf & _
"        var chart = am4core.create(chartId, am4charts.XYChart);" & vbCrLf & _
"        chart.colors.step = 2;" & vbCrLf & _
"" & vbCrLf & _
"        chart.cursor = new am4charts.XYCursor();" & vbCrLf & _
"        chart.cursor.behavior = ""zoomXY"";" & vbCrLf & _
"" & vbCrLf & _
"        chart.legend = new am4charts.Legend();" & vbCrLf & _
"        chart.legend.position = ""top"";" & vbCrLf & _
"        chart.legend.paddingBottom = 20;" & vbCrLf & _
"        chart.legend.labels.template.maxWidth = 95;" & vbCrLf & _
"" & vbCrLf & _
"        var xAxis = chart.xAxes.push(new am4charts.CategoryAxis());" & vbCrLf & _
"        xAxis.dataFields.category = ""type"";" & vbCrLf & _
"        xAxis.renderer.cellStartLocation = 0.1;" & vbCrLf & _
"        xAxis.renderer.cellEndLocation = 0.9;" & vbCrLf & _
"        xAxis.renderer.grid.template.location = 0;" & vbCrLf & _
"" & vbCrLf & _
"        function getExtremes(data) {" & vbCrLf & _
"          const initialMax = -Infinity;" & vbCrLf & _
"          const initialMin = Infinity;" & vbCrLf & _
"" & vbCrLf & _
"          const extremes = data.reduce(" & vbCrLf & _
"            (acc, item) => {" & vbCrLf & _
"              const value = parseFloat(item.total);" & vbCrLf & _
"" & vbCrLf & _
"              if (value > acc.max) {" & vbCrLf & _
"                acc.max = value;" & vbCrLf & _
"              }" & vbCrLf & _
"              if (value < acc.min) {" & vbCrLf & _
"                acc.min = value;" & vbCrLf & _
"              }" & vbCrLf & _
"" & vbCrLf & _
"              return acc;" & vbCrLf & _
"            }," & vbCrLf & _
"            { max: initialMax, min: initialMin }" & vbCrLf & _
"          );" & vbCrLf & _
"" & vbCrLf & _
"          return extremes;" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        const { max: highestValue, min: lowestValue } = getExtremes(data);" & vbCrLf & _
"" & vbCrLf & _
"        var yAxis = chart.yAxes.push(new am4charts.ValueAxis());" & vbCrLf & _
"        yAxis.min = chartId == ""rmpnmibase"" ? -50000 : lowestValue;" & vbCrLf & _
"        yAxis.max = chartId == ""rmpnmibase"" ? 50000 : highestValue;" & vbCrLf & _
"        yAxis.renderer.minGridDistance = 20;" & vbCrLf & _
"" & vbCrLf & _
"        function createSeries(value, name) {" & vbCrLf & _
"          var series = chart.series.push(new am4charts.ColumnSeries());" & vbCrLf & _
"          series.dataFields.valueY = value;" & vbCrLf & _
"          series.dataFields.categoryX = ""type"";" & vbCrLf & _
"          series.name = name;" & vbCrLf & _
"          if (chartId == ""rmpnmibase"") {" & vbCrLf & _
"            series.columns.template.tooltipText = ""{name}: [bold]{valueY}[/]"";" & vbCrLf & _
"          } else {" & vbCrLf & _
"            series.columns.template.tooltipText =" & vbCrLf & _
"              ""{categoryX}: [bold]{valueY}[/]"";" & vbCrLf & _
"          }" & vbCrLf & _
"" & vbCrLf & _
"          series.columns.template.fontSize = 14;" & vbCrLf & _
"" & vbCrLf & _
"          series.events.on(""hidden"", arrangeColumns);" & vbCrLf & _
"          series.events.on(""shown"", arrangeColumns);" & vbCrLf & _
"" & vbCrLf & _
"          var bullet = series.bullets.push(new am4charts.LabelBullet());" & vbCrLf & _
"          bullet.interactionsEnabled = true;" & vbCrLf & _
"          bullet.dy = 30;" & vbCrLf & _
"          bullet.label.fill = am4core.color(""#ffffff"");" & vbCrLf & _
"" & vbCrLf & _
"          return series;" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        const categories = [...new Set(data.map((item) => item.type))];" & vbCrLf & _
"        const seriesNames = [...new Set(data.map((item) => item.margin))];" & vbCrLf & _
"" & vbCrLf & _
"        const chartData = categories.map((category) => {" & vbCrLf & _
"          const dataPoint = { type: category };" & vbCrLf & _
"          seriesNames.forEach((name) => {" & vbCrLf & _
"            const entry = data.find(" & vbCrLf & _
"              (d) => d.type === category && d.margin === name" & vbCrLf & _
"            );" & vbCrLf & _
"            dataPoint[name] = entry ? parseFloat(entry.total) : 0;" & vbCrLf & _
"          });" & vbCrLf & _
"          return dataPoint;" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        chart.data = chartData;" & vbCrLf & _
"" & vbCrLf & _
"        seriesNames.forEach((name) => {" & vbCrLf & _
"          createSeries(name, name);" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        function arrangeColumns() {" & vbCrLf & _
"          var series = chart.series.getIndex(0);" & vbCrLf & _
"" & vbCrLf & _
"          var w =" & vbCrLf & _
"            1 -" & vbCrLf & _
"            xAxis.renderer.cellStartLocation -" & vbCrLf & _
"            (1 - xAxis.renderer.cellEndLocation);" & vbCrLf & _
"          if (series.dataItems.length > 1) {" & vbCrLf & _
"            var x0 = xAxis.getX(series.dataItems.getIndex(0), ""categoryX"");" & vbCrLf & _
"            var x1 = xAxis.getX(series.dataItems.getIndex(1), ""categoryX"");" & vbCrLf & _
"            var delta = ((x1 - x0) / chart.series.length) * w;" & vbCrLf & _
"            if (am4core.isNumber(delta)) {" & vbCrLf & _
"              var middle = chart.series.length / 2;" & vbCrLf & _
"" & vbCrLf & _
"              var newIndex = 0;" & vbCrLf & _
"              chart.series.each(function (series) {" & vbCrLf & _
"                if (!series.isHidden && !series.isHiding) {" & vbCrLf & _
"                  series.dummyData = newIndex;" & vbCrLf & _
"                  newIndex++;" & vbCrLf & _
"                } else {" & vbCrLf & _
"                  series.dummyData = chart.series.indexOf(series);" & vbCrLf & _
"                }" & vbCrLf & _
"              });" & vbCrLf & _
"              var visibleCount = newIndex;" & vbCrLf & _
"              var newMiddle = visibleCount / 2;" & vbCrLf & _
"" & vbCrLf & _
"              chart.series.each(function (series) {" & vbCrLf & _
"                var trueIndex = chart.series.indexOf(series);" & vbCrLf & _
"                var newIndex = series.dummyData;" & vbCrLf & _
"" & vbCrLf & _
"                var dx = (newIndex - trueIndex + middle - newMiddle) * delta;" & vbCrLf & _
"" & vbCrLf & _
"                series.animate(" & vbCrLf & _
"                  { property: ""dx"", to: dx }," & vbCrLf & _
"                  series.interpolationDuration," & vbCrLf & _
"                  series.interpolationEasing" & vbCrLf & _
"                );" & vbCrLf & _
"                series.bulletsContainer.animate(" & vbCrLf & _
"                  { property: ""dx"", to: dx }," & vbCrLf & _
"                  series.interpolationDuration," & vbCrLf & _
"                  series.interpolationEasing" & vbCrLf & _
"                );" & vbCrLf & _
"              });" & vbCrLf & _
"            }" & vbCrLf & _
"          }" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        chart.exporting.menu = new am4core.ExportMenu();" & vbCrLf & _
"      };" & vbCrLf & _
"      const populateTable2Data = (data, id) => {" & vbCrLf & _
"        const headerRow = document.querySelector(""#"" + id + "" #headerRow"");" & vbCrLf & _
"        const tbody = document.querySelector(""#"" + id + "" tbody"");" & vbCrLf & _
"" & vbCrLf & _
"        const marginTypes = new Set();" & vbCrLf & _
"        const typesOfMargins = new Set();" & vbCrLf & _
"" & vbCrLf & _
"        data.forEach((item) => {" & vbCrLf & _
"          marginTypes.add(item.type);" & vbCrLf & _
"          typesOfMargins.add(item.margin);" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        headerRow.innerHTML =" & vbCrLf & _
"          ""<th></th>"" +" & vbCrLf & _
"          Array.from(marginTypes)" & vbCrLf & _
"            .map(" & vbCrLf & _
"              (type) =>" & vbCrLf & _
"                ""<th>"" +" & vbCrLf & _
"                type.replace(/([A-Z][a-z]+|[A-Z]+(?![a-z]))/g, "" $1"").trim() +" & vbCrLf & _
"                ""</th>""" & vbCrLf & _
"            )" & vbCrLf & _
"            .join("""");" & vbCrLf & _
"" & vbCrLf & _
"        const totalsMap = {};" & vbCrLf & _
"" & vbCrLf & _
"        typesOfMargins.forEach((marginType) => {" & vbCrLf & _
"          totalsMap[marginType] = {};" & vbCrLf & _
"          marginTypes.forEach((type) => {" & vbCrLf & _
"            totalsMap[marginType][type] = 0;" & vbCrLf & _
"          });" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        data.forEach((item) => {" & vbCrLf & _
"          if (totalsMap[item.margin]) {" & vbCrLf & _
"            totalsMap[item.margin][item.type] += parseFloat(item.total);" & vbCrLf & _
"          }" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        Object.keys(totalsMap).forEach((marginType) => {" & vbCrLf & _
"          const row = document.createElement(""tr"");" & vbCrLf & _
"          const marginTypeCell = document.createElement(""td"");" & vbCrLf & _
"          marginTypeCell.textContent = marginType;" & vbCrLf & _
"          row.appendChild(marginTypeCell);" & vbCrLf & _
"" & vbCrLf & _
"          marginTypes.forEach((type) => {" & vbCrLf & _
"            const td = document.createElement(""td"");" & vbCrLf & _
"            td.textContent = totalsMap[marginType][type].toFixed(2);" & vbCrLf & _
"            row.appendChild(td);" & vbCrLf & _
"          });" & vbCrLf & _
"" & vbCrLf & _
"          tbody.appendChild(row);" & vbCrLf & _
"        });" & vbCrLf & _
"      };" & vbCrLf & _
"      const populateTable3Data = (data, id) => {" & vbCrLf & _
"        const headerRow = document.querySelector(""#"" + id + "" #headerRow"");" & vbCrLf & _
"        const tbody = document.querySelector(""#"" + id + "" tbody"");" & vbCrLf & _
"" & vbCrLf & _
"        headerRow.innerHTML = """";" & vbCrLf & _
"        tbody.innerHTML = """";" & vbCrLf & _
"" & vbCrLf & _
"        headerRow.innerHTML = ""<th>Margin</th><th>Total</th>"";" & vbCrLf & _
"" & vbCrLf & _
"        const totalsMap = {};" & vbCrLf & _
"" & vbCrLf & _
"        data.forEach((item) => {" & vbCrLf & _
"          if (!totalsMap[item.type]) {" & vbCrLf & _
"            totalsMap[item.type] = 0;" & vbCrLf & _
"          }" & vbCrLf & _
"          totalsMap[item.type] += parseFloat(item.total);" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        Object.keys(totalsMap).forEach((type) => {" & vbCrLf & _
"          const row = document.createElement(""tr"");" & vbCrLf & _
"" & vbCrLf & _
"          const typeCell = document.createElement(""td"");" & vbCrLf & _
"          typeCell.textContent = type;" & vbCrLf & _
"          row.appendChild(typeCell);" & vbCrLf & _
"" & vbCrLf & _
"          const totalCell = document.createElement(""td"");" & vbCrLf & _
"          totalCell.textContent = totalsMap[type].toFixed(2);" & vbCrLf & _
"          row.appendChild(totalCell);" & vbCrLf & _
"" & vbCrLf & _
"          tbody.appendChild(row);" & vbCrLf & _
"        });" & vbCrLf & _
"      };" & vbCrLf & _
"      const getUniqueValue = (data) => {" & vbCrLf & _
"        const uniqueTypes = [" & vbCrLf & _
"          ...new Set(" & vbCrLf & _
"            data" & vbCrLf & _
"              .map((item) => item.type)" & vbCrLf & _
"              .filter((type) => type && type.trim() !== ""(blank)"")" & vbCrLf & _
"          )," & vbCrLf & _
"        ];" & vbCrLf & _
"        const selectAllOption = document.createElement(""option"");" & vbCrLf & _
"        selectAllOption.value = ""selectAll"";" & vbCrLf & _
"        selectAllOption.text = ""Select All"";" & vbCrLf & _
"        select.insertBefore(selectAllOption, select.firstChild);" & vbCrLf & _
"" & vbCrLf & _
"        uniqueTypes.forEach((type) => {" & vbCrLf & _
"          const option = document.createElement(""option"");" & vbCrLf & _
"          option.value = type;" & vbCrLf & _
"          option.text = type;" & vbCrLf & _
"          select.appendChild(option);" & vbCrLf & _
"        });" & vbCrLf & _
"      };" & vbCrLf & _
"      const totalCost = (data) => {" & vbCrLf & _
"        const totals = {};" & vbCrLf & _
"" & vbCrLf & _
"        data.forEach((item) => {" & vbCrLf & _
"          if (!totals[item.type]) {" & vbCrLf & _
"            totals[item.type] = 0;" & vbCrLf & _
"          }" & vbCrLf & _
"          totals[item.type] += parseFloat(item.total);" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        return Object.keys(totals).map((type) => ({" & vbCrLf & _
"          type: type," & vbCrLf & _
"          total: totals[type].toFixed(2)," & vbCrLf & _
"        }));" & vbCrLf & _
"      };" & vbCrLf & _
"      const fyActualMarginVsPOE = (dataArr, chartId) => {" & vbCrLf & _
"        am4core.useTheme(am4themes_animated);" & vbCrLf & _
"" & vbCrLf & _
"        var chart = am4core.create(chartId, am4charts.XYChart);" & vbCrLf & _
"" & vbCrLf & _
"        const transformedData = [];" & vbCrLf & _
"" & vbCrLf & _
"        dataArr.forEach((margin) => {" & vbCrLf & _
"          margin.totals.forEach((total) => {" & vbCrLf & _
"            let existingEntry = transformedData.find(" & vbCrLf & _
"              (entry) => entry.type === total.type" & vbCrLf & _
"            );" & vbCrLf & _
"            if (!existingEntry) {" & vbCrLf & _
"              existingEntry = { type: total.type };" & vbCrLf & _
"              transformedData.push(existingEntry);" & vbCrLf & _
"            }" & vbCrLf & _
"" & vbCrLf & _
"            existingEntry[margin.typesofmargin] = total.total;" & vbCrLf & _
"          });" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        chart.data = transformedData;" & vbCrLf & _
"" & vbCrLf & _
"        var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());" & vbCrLf & _
"        categoryAxis.dataFields.category = ""type"";" & vbCrLf & _
"        categoryAxis.renderer.labels.template.fontSize = 14;" & vbCrLf & _
"        categoryAxis.renderer.grid.template.location = 0;" & vbCrLf & _
"        categoryAxis.renderer.minGridDistance = 20;" & vbCrLf & _
"        categoryAxis.title.text = ""Type"";" & vbCrLf & _
"" & vbCrLf & _
"        var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());" & vbCrLf & _
"        valueAxis.renderer.minWidth = 35;" & vbCrLf & _
"        valueAxis.title.text = ""Total"";" & vbCrLf & _
"" & vbCrLf & _
"        dataArr.forEach((margin) => {" & vbCrLf & _
"          var series = chart.series.push(new am4charts.LineSeries());" & vbCrLf & _
"          series.dataFields.valueY = margin.typesofmargin;" & vbCrLf & _
"          series.dataFields.categoryX = ""type"";" & vbCrLf & _
"          series.name = margin.typesofmargin;" & vbCrLf & _
"          series.strokeWidth = 2;" & vbCrLf & _
"          series.tooltipText = ""{name}: [bold]{valueY}[/]"";" & vbCrLf & _
"" & vbCrLf & _
"          var bullet = series.bullets.push(new am4charts.CircleBullet());" & vbCrLf & _
"          bullet.circle.fill = am4core.color(""#FF5733"");" & vbCrLf & _
"          bullet.circle.stroke = am4core.color(""#fff"");" & vbCrLf & _
"          bullet.circle.strokeWidth = 2;" & vbCrLf & _
"          bullet.circle.radius = 6;" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        chart.cursor = new am4charts.XYCursor();" & vbCrLf & _
"        chart.cursor.behavior = ""zoomY"";" & vbCrLf & _
"" & vbCrLf & _
"        chart.legend = new am4charts.Legend();" & vbCrLf & _
"" & vbCrLf & _
"        chart.scrollbarX = new am4core.Scrollbar();" & vbCrLf & _
"        chart.exporting.menu = new am4core.ExportMenu();" & vbCrLf & _
"      };" & vbCrLf & _
"      const nmiClusteredChart = (data, chartId) => {" & vbCrLf & _
"        loadingElement.style.display = ""block"";" & vbCrLf & _
"        if (nmiChartVar) {" & vbCrLf & _
"          nmiChartVar.dispose();" & vbCrLf & _
"        }" & vbCrLf & _
"        am4core.useTheme(am4themes_animated);" & vbCrLf & _
"        nmiChartVar = am4core.create(chartId, am4charts.XYChart);" & vbCrLf & _
"        nmiChartVar.colors.step = 2;" & vbCrLf & _
"" & vbCrLf & _
"        nmiChartVar.cursor = new am4charts.XYCursor();" & vbCrLf & _
"        nmiChartVar.cursor.behavior = ""zoomXY"";" & vbCrLf & _
"" & vbCrLf & _
"        nmiChartVar.legend = new am4charts.Legend();" & vbCrLf & _
"        nmiChartVar.legend.position = ""top"";" & vbCrLf & _
"        nmiChartVar.legend.paddingBottom = 20;" & vbCrLf & _
"        nmiChartVar.legend.labels.template.maxWidth = 95;" & vbCrLf & _
"" & vbCrLf & _
"        var xAxis = nmiChartVar.xAxes.push(new am4charts.CategoryAxis());" & vbCrLf & _
"        xAxis.dataFields.category = ""type"";" & vbCrLf & _
"        xAxis.renderer.cellStartLocation = 0.1;" & vbCrLf & _
"        xAxis.renderer.cellEndLocation = 0.9;" & vbCrLf & _
"        xAxis.renderer.grid.template.location = 0;" & vbCrLf & _
"" & vbCrLf & _
"        function getExtremes(data) {" & vbCrLf & _
"          const initialMax = -Infinity;" & vbCrLf & _
"          const initialMin = Infinity;" & vbCrLf & _
"" & vbCrLf & _
"          const extremes = data.reduce(" & vbCrLf & _
"            (acc, item) => {" & vbCrLf & _
"              const value = parseFloat(item.total);" & vbCrLf & _
"" & vbCrLf & _
"              if (value > acc.max) {" & vbCrLf & _
"                acc.max = value;" & vbCrLf & _
"              }" & vbCrLf & _
"              if (value < acc.min) {" & vbCrLf & _
"                acc.min = value;" & vbCrLf & _
"              }" & vbCrLf & _
"" & vbCrLf & _
"              return acc;" & vbCrLf & _
"            }," & vbCrLf & _
"            { max: initialMax, min: initialMin }" & vbCrLf & _
"          );" & vbCrLf & _
"" & vbCrLf & _
"          return extremes;" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        const { max: highestValue, min: lowestValue } = getExtremes(data);" & vbCrLf & _
"" & vbCrLf & _
"        var yAxis = nmiChartVar.yAxes.push(new am4charts.ValueAxis());" & vbCrLf & _
"        yAxis.min = chartId == ""rmpnmibase"" ? -50000 : lowestValue;" & vbCrLf & _
"        yAxis.max = chartId == ""rmpnmibase"" ? 50000 : highestValue;" & vbCrLf & _
"        yAxis.renderer.minGridDistance = 20;" & vbCrLf & _
"" & vbCrLf & _
"        function createSeries(value, name) {" & vbCrLf & _
"          var series = nmiChartVar.series.push(new am4charts.ColumnSeries());" & vbCrLf & _
"          series.dataFields.valueY = value;" & vbCrLf & _
"          series.dataFields.categoryX = ""type"";" & vbCrLf & _
"          series.name = name;" & vbCrLf & _
"          series.columns.template.tooltipText = ""{name}: [bold]{valueY}[/]"";" & vbCrLf & _
"          series.columns.template.fontSize = 14;" & vbCrLf & _
"" & vbCrLf & _
"          series.events.on(""hidden"", arrangeColumns);" & vbCrLf & _
"          series.events.on(""shown"", arrangeColumns);" & vbCrLf & _
"" & vbCrLf & _
"          var bullet = series.bullets.push(new am4charts.LabelBullet());" & vbCrLf & _
"          bullet.interactionsEnabled = true;" & vbCrLf & _
"          bullet.dy = 30;" & vbCrLf & _
"          bullet.label.fill = am4core.color(""#ffffff"");" & vbCrLf & _
"" & vbCrLf & _
"          return series;" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        const categories = [...new Set(data.map((item) => item.type))];" & vbCrLf & _
"        const seriesNames = [...new Set(data.map((item) => item.margin))];" & vbCrLf & _
"" & vbCrLf & _
"        const chartData = categories.map((category) => {" & vbCrLf & _
"          const dataPoint = { type: category };" & vbCrLf & _
"          seriesNames.forEach((name) => {" & vbCrLf & _
"            const entry = data.find(" & vbCrLf & _
"              (d) => d.type === category && d.margin === name" & vbCrLf & _
"            );" & vbCrLf & _
"            dataPoint[name] = entry ? parseFloat(entry.total) : 0;" & vbCrLf & _
"          });" & vbCrLf & _
"          return dataPoint;" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        nmiChartVar.data = chartData;" & vbCrLf & _
"" & vbCrLf & _
"        seriesNames.forEach((name) => {" & vbCrLf & _
"          createSeries(name, name);" & vbCrLf & _
"        });" & vbCrLf & _
"" & vbCrLf & _
"        function arrangeColumns() {" & vbCrLf & _
"          var series = nmiChartVar.series.getIndex(0);" & vbCrLf & _
"" & vbCrLf & _
"          var w =" & vbCrLf & _
"            1 -" & vbCrLf & _
"            xAxis.renderer.cellStartLocation -" & vbCrLf & _
"            (1 - xAxis.renderer.cellEndLocation);" & vbCrLf & _
"          if (series.dataItems.length > 1) {" & vbCrLf & _
"            var x0 = xAxis.getX(series.dataItems.getIndex(0), ""categoryX"");" & vbCrLf & _
"            var x1 = xAxis.getX(series.dataItems.getIndex(1), ""categoryX"");" & vbCrLf & _
"            var delta = ((x1 - x0) / nmiChartVar.series.length) * w;" & vbCrLf & _
"            if (am4core.isNumber(delta)) {" & vbCrLf & _
"              var middle = nmiChartVar.series.length / 2;" & vbCrLf & _
"" & vbCrLf & _
"              var newIndex = 0;" & vbCrLf & _
"              nmiChartVar.series.each(function (series) {" & vbCrLf & _
"                if (!series.isHidden && !series.isHiding) {" & vbCrLf & _
"                  series.dummyData = newIndex;" & vbCrLf & _
"                  newIndex++;" & vbCrLf & _
"                } else {" & vbCrLf & _
"                  series.dummyData = nmiChartVar.series.indexOf(series);" & vbCrLf & _
"                }" & vbCrLf & _
"              });" & vbCrLf & _
"              var visibleCount = newIndex;" & vbCrLf & _
"              var newMiddle = visibleCount / 2;" & vbCrLf & _
"" & vbCrLf & _
"              nmiChartVar.series.each(function (series) {" & vbCrLf & _
"                var trueIndex = nmiChartVar.series.indexOf(series);" & vbCrLf & _
"                var newIndex = series.dummyData;" & vbCrLf & _
"" & vbCrLf & _
"                var dx = (newIndex - trueIndex + middle - newMiddle) * delta;" & vbCrLf & _
"" & vbCrLf & _
"                series.animate(" & vbCrLf & _
"                  { property: ""dx"", to: dx }," & vbCrLf & _
"                  series.interpolationDuration," & vbCrLf & _
"                  series.interpolationEasing" & vbCrLf & _
"                );" & vbCrLf & _
"                series.bulletsContainer.animate(" & vbCrLf & _
"                  { property: ""dx"", to: dx }," & vbCrLf & _
"                  series.interpolationDuration," & vbCrLf & _
"                  series.interpolationEasing" & vbCrLf & _
"                );" & vbCrLf & _
"              });" & vbCrLf & _
"            }" & vbCrLf & _
"          }" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        nmiChartVar.exporting.menu = new am4core.ExportMenu();" & vbCrLf & _
"        nmiChartVar.events.on(""validated"", function () {" & vbCrLf & _
"          loadingElement.style.display = ""none"";" & vbCrLf & _
"        });" & vbCrLf & _
"        return nmiChartVar;" & vbCrLf & _
"      };" & vbCrLf & _
"    </script>" & vbCrLf & _
"    <!-- variables and execution -->" & vbCrLf & _
"    <script>" & vbCrLf & _
"      var combinedJsonData = {{combinedJson}};" & vbCrLf & _
"      var summaryJsonData = {{summaryJson}};" & vbCrLf & _
"      var versionData = {{versionData}};" & vbCrLf & _
"      var nmiJsonData = {{nmiJson}}" & vbCrLf & _
"" & vbCrLf & _
"      var nmiChartVar;" & vbCrLf & _
"      const typeList =[" & vbCrLf & _
"            ""RetailMargin""," & vbCrLf & _
"            ""Revenue""," & vbCrLf & _
"            ""Network""," & vbCrLf & _
"            ""Capacity""," & vbCrLf & _
"            ""WholesaleEnergy""," & vbCrLf & _
"            ""MarketFees""," & vbCrLf & _
"            ""ESS""," & vbCrLf & _
"            ""LGC""," & vbCrLf & _
"            ""STC""," & vbCrLf & _
"            ""Commission""" & vbCrLf & _
"        ]" & vbCrLf & _
"" & vbCrLf & _
"      const achievedMargin2024 = transformDataForChart(combinedJsonData, ""Sum of Achieved Margin FY2024"")" & vbCrLf & _
"      const dataFy2024 =  transformDataForChart(combinedJsonData, ""FY2024"")" & vbCrLf & _
"      const dataFy2025 =  transformDataForCluster(combinedJsonData, ""FY2025"")" & vbCrLf & _
"      const dataFy2026 = transformDataForCluster(combinedJsonData, ""FY2026"")" & vbCrLf & _
"      const dataRmperNMIBase = transformDataForCluster(nmiJsonData.filter(data => data.typesofmargin != ""(blank)"" && data.typesofmargin != ""Grand Total""));" & vbCrLf & _
"      var updatedDataRmperNMIBase = []" & vbCrLf & _
"      var loadingElement = document.getElementById('loading');" & vbCrLf & _
"      const select = document.getElementById('nmi');" & vbCrLf & _
"" & vbCrLf & _
"      xyChart(achievedMargin2024, ""hpfy24a"", ""Historical Performance FY2024 Actuals"");" & vbCrLf & _
"      animatedPieChart(achievedMargin2024, ""pfify24"");" & vbCrLf & _
"      populateTable1Data(achievedMargin2024);" & vbCrLf & _
"      clusteredChart(dataFy2025, ""fy25avspoe"");" & vbCrLf & _
"      populateTable2Data(dataFy2025, 'table2');" & vbCrLf & _
"      clusteredChart(dataFy2026, ""fy26avspoe"");" & vbCrLf & _
"      populateTable2Data(dataFy2026, 'table3');" & vbCrLf & _
"      animatedPieChart(totalCost(transformDataForCluster(summaryJsonData)), ""acap"");" & vbCrLf & _
"      populateTable3Data(totalCost(transformDataForCluster(summaryJsonData)), 'table5');" & vbCrLf & _
"      // Example usage" & vbCrLf & _
"      const combinedSummary = []" & vbCrLf & _
"      const prepareDateForChart = () => {" & vbCrLf & _
"        function getLast6IfStartsWithFY(inputString) {" & vbCrLf & _
"          if (inputString.length < 6) {" & vbCrLf & _
"            return null;" & vbCrLf & _
"          }" & vbCrLf & _
"" & vbCrLf & _
"          // Extract the last 6 characters" & vbCrLf & _
"          const last6 = inputString.slice(-6);" & vbCrLf & _
"" & vbCrLf & _
"          if (last6.startsWith('FY')) {" & vbCrLf & _
"            return last6;" & vbCrLf & _
"          }" & vbCrLf & _
"        }" & vbCrLf & _
"" & vbCrLf & _
"        summaryJsonData.forEach((data, index) => {" & vbCrLf & _
"          const dataOb = {" & vbCrLf & _
"            typesofmargin: data.typesofmargin," & vbCrLf & _
"            totals: transformDataForChart(summaryJsonData, data.typesofmargin)" & vbCrLf & _
"          }" & vbCrLf & _
"          if (data.typesofmargin != ""Grand Total"")" & vbCrLf & _
"            combinedSummary.push(dataOb)" & vbCrLf & _
"        })" & vbCrLf & _
"        combinedSummary.sort((a, b) => {" & vbCrLf & _
"            if (a.typesofmargin < b.typesofmargin) return -1;" & vbCrLf & _
"            if (a.typesofmargin > b.typesofmargin) return 1;" & vbCrLf & _
"            return 0;" & vbCrLf & _
"          });" & vbCrLf & _
"      }" & vbCrLf & _
"      prepareDateForChart()" & vbCrLf & _
"" & vbCrLf & _
"      fyActualMarginVsPOE(combinedSummary, ""fyamvspoe"");" & vbCrLf & _
"      populateTable2Data(transformDataForCluster(combinedJsonData), 'table4');" & vbCrLf & _
"      nmiClusteredChart(dataRmperNMIBase, ""rmpnmibase"");" & vbCrLf & _
"      getUniqueValue(dataRmperNMIBase)" & vbCrLf & _
"    </script>" & vbCrLf & _
"    <!-- footer -->" & vbCrLf & _
"    <script>" & vbCrLf & _
"      const footerContainer = document.querySelector("".footer-container"");" & vbCrLf & _
"" & vbCrLf & _
"      versionData.forEach((item) => {" & vbCrLf & _
"        const footerItem = document.createElement(""div"");" & vbCrLf & _
"        footerItem.className = ""footer-item"";" & vbCrLf & _
"        footerItem.innerHTML =""<h3>"" + item.version + ""</h3>"" +" & vbCrLf & _
"          ""<p>"" + item.effectiveDate + ""</p>"";" & vbCrLf & _
"        footerContainer.appendChild(footerItem);" & vbCrLf & _
"      });" & vbCrLf & _
"    </script>" & vbCrLf & _
"    <!-- event listener -->" & vbCrLf & _
"    <script>" & vbCrLf & _
"      function reloadChart(newData) {" & vbCrLf & _
"        nmiClusteredChart(newData, ""rmpnmibase"");" & vbCrLf & _
"        isChartLoading = true;" & vbCrLf & _
"      }" & vbCrLf & _
"      select.addEventListener(""change"", function () {" & vbCrLf & _
"        if (select.value === ""selectAll"") {" & vbCrLf & _
"          updatedDataRmperNMIBase = dataRmperNMIBase;" & vbCrLf & _
"          for (let i = 0; i < select.options.length; i++) {" & vbCrLf & _
"            if (select.options[i].value !== ""selectAll"") {" & vbCrLf & _
"              select.options[i].selected = true; // Select all other options" & vbCrLf & _
"            }" & vbCrLf & _
"          }" & vbCrLf & _
"        } else {" & vbCrLf & _
"          updatedDataRmperNMIBase = transformDataForCluster(" & vbCrLf & _
"            nmiJsonData.filter(" & vbCrLf & _
"              (data) =>" & vbCrLf & _
"                data.typesofmargin != ""(blank)"" &&" & vbCrLf & _
"                data.typesofmargin != ""Grand Total""" & vbCrLf & _
"            )" & vbCrLf & _
"          ).filter((data) => data.type == select.value);" & vbCrLf & _
"        }" & vbCrLf & _
"        reloadChart(updatedDataRmperNMIBase);" & vbCrLf & _
"        console.log(""dataRmperNMIBase"", updatedDataRmperNMIBase, select.value);" & vbCrLf & _
"      });" & vbCrLf & _
"    </script>" & vbCrLf & _
"  </body>" & vbCrLf & _
"</html>" & vbCrLf & _
"" & vbCrLf & _
""