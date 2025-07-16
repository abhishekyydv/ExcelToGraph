let currentCharts = [];

document.getElementById('file-input').addEventListener('change', function (event) {
    const file = event.target.files[0];
    if (!file) return;

    const chartsContainer = document.getElementById('charts-container');
    chartsContainer.innerHTML = '';
    currentCharts = [];

    if (file.name.endsWith('.csv')) {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: results => renderSheet(results.data, 'CSV Data')
        });
    } else if (file.name.endsWith('.xlsx')) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const workbook = XLSX.read(e.target.result, { type: 'array' });

            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

                const headers = jsonData[0];
                const rows = jsonData.slice(1).map(row => {
                    const obj = {};
                    headers.forEach((h, i) => {
                        let value = row[i];

                        // Handle Excel date serials
                        if (typeof value === 'number' && value > 30000 && value < 60000) {
                            const parsedDate = XLSX.SSF.parse_date_code(value);
                            if (parsedDate) {
                                const pad = n => n.toString().padStart(2, '0');
                                value = `${pad(parsedDate.d)}-${pad(parsedDate.m)}-${parsedDate.y} ${pad(parsedDate.H)}:${pad(parsedDate.M)}:${pad(parsedDate.S)}`;
                            }
                        }

                        obj[h] = value;
                    });
                    return obj;
                });

                renderSheet(rows, sheetName);
            });
        };
        reader.readAsArrayBuffer(file);
    }
});

function renderSheet(data, sheetName) {
    if (!data.length) {
        alert(`Sheet "${sheetName}" is empty or invalid.`);
        return;
    }

    const containerId = `chart-${sheetName.replace(/\s+/g, '-')}`;
    const tableId = `table-${sheetName.replace(/\s+/g, '-')}`;
    const selectId = `select-${sheetName.replace(/\s+/g, '-')}`;

    const section = document.createElement('section');
    section.innerHTML = `
        <h2>${sheetName}</h2>
        <label for="${selectId}">Select X-axis:</label>
        <select id="${selectId}"></select>
        <div id="${containerId}" class="chart-container"></div>
        <section id="excel-table-container-${tableId}" class="excel-table-container hidden">
            <h3>Data Table View (${sheetName})</h3>
            <table id="${tableId}" class="styled-table"></table>
        </section>
    `;

    document.getElementById('charts-container').appendChild(section);

    const headers = Object.keys(data[0]);
    const xAxisSelect = document.getElementById(selectId);
    xAxisSelect.innerHTML = '';

    headers.forEach(header => {
        const opt = document.createElement('option');
        opt.value = header;
        opt.text = header;
        xAxisSelect.appendChild(opt);
    });

    const drawChart = xKey => {
        renderChart(data, xKey, containerId, tableId);
    };

    xAxisSelect.onchange = () => drawChart(xAxisSelect.value);
    drawChart(headers[0]);
}

function renderChart(data, xKey, containerId, tableId) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';

    const xAxisLabels = data.map(row => row[xKey]);
    const headers = Object.keys(data[0]);

    const series = headers.filter(h => h !== xKey).map(h => ({
        name: h,
        data: data.map(row => parseFloat(row[h]) || null)
    }));

    const chart = Highcharts.chart(containerId, {
        chart: {
            zoomType: 'xy',
            events: {
                load: function () {
                    const chart = this;
                    Highcharts.addEvent(chart.container, 'wheel', function (e) {
                        e.preventDefault();
                        const axis = chart.xAxis[0];
                        const { min, max } = axis.getExtremes();
                        const range = max - min;
                        const center = min + range / 2;
                        const zoom = e.deltaY > 0 ? 1.2 : 0.8;
                        axis.setExtremes(center - (range * zoom / 2), center + (range * zoom / 2));
                    });
                }
            }
        },
        title: {
            text: 'Visualized Data',
            align: 'left'
        },
        xAxis: {
            categories: xAxisLabels,
            title: { text: xKey },
            labels: {
                rotation: -45,
                style: {
                    fontSize: '10px'
                }
            }
        },
        yAxis: { title: { text: 'Values' } },
        legend: {
            layout: 'vertical',
            align: 'right',
            verticalAlign: 'middle'
        },
        tooltip: { shared: true },
        series: series,
        exporting: {
            enabled: true,
            buttons: {
                contextButton: {
                    menuItems: [
                        {
                            text: 'View Fullscreen',
                            onclick: () => toggleFullscreen(containerId, chart)
                        },
                        {
                            text: 'Print Chart',
                            onclick: () => chart.print()
                        },
                        'downloadPDF',
                        'downloadCSV',
                        'downloadXLS',
                        {
                            text: 'Toggle Table',
                            onclick: () => toggleTableView(tableId)
                        },
                        {
                            text: 'Reset Zoom',
                            onclick: () => resetZoom(chart)
                        }
                    ]
                }
            }
        },
        credits: { enabled: false },
        responsive: {
            rules: [{
                condition: { maxWidth: 600 },
                chartOptions: {
                    legend: {
                        layout: 'horizontal',
                        align: 'center',
                        verticalAlign: 'bottom'
                    }
                }
            }]
        }
    });

    renderExcelTable(data, tableId);
    currentCharts.push(chart);
}

function renderExcelTable(data, tableId) {
    const table = document.getElementById(tableId);
    table.innerHTML = '';

    const headers = Object.keys(data[0]);
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        tr.appendChild(th);
    });
    thead.appendChild(tr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    data.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
            const td = document.createElement('td');
            td.textContent = row[h] ?? '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
}

function toggleTableView(tableId) {
    const container = document.getElementById(`excel-table-container-${tableId}`);
    container.classList.toggle('hidden');
}

function toggleFullscreen(containerId, chart) {
    const el = document.getElementById(containerId);
    if (!document.fullscreenElement) {
        el.requestFullscreen();
    } else {
        document.exitFullscreen();
    }
    document.addEventListener('fullscreenchange', () => {
        if (chart) chart.reflow();
    });
}

function resetZoom(chart) {
    chart.xAxis[0].setExtremes(null, null);
    chart.yAxis[0].setExtremes(null, null);
}
