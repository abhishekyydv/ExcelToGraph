document.getElementById('excel-file').addEventListener('change', function (event) {
    const file = event.target.files[0];

    if (file) {
        const reader = new FileReader();

        reader.onload = function (e) {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'array' });

            // Assuming the first sheet is the one to be used
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Extract the headings (first row) and the data (rest of the rows)
            const headings = jsonData[0]; // All headings from the first row
            const dataRows = jsonData.slice(1); // Data rows excluding the first row

            // Prepare the X-axis data (Serial No.)
            const xAxisData = dataRows.map((row, index) => index + 1); // Serial numbers (1, 2, 3, ...)

            // Prepare the Highcharts series dynamically
            const series = headings.map((heading, index) => ({
                name: heading,
                data: dataRows.map(row => row[index] || null) // Data for each column
            }));

            // Create the chart with the dynamic data
            Highcharts.chart('container', {
                title: {
                    text: 'Graph based on data from your Excel Sheet',
                    align: 'left'
                },

                subtitle: {
                    text: 'Source: Excel file uploaded by you',
                    align: 'left'
                },

                yAxis: {
                    title: {
                        text: 'Values'
                    }
                },

                xAxis: {
                    categories: xAxisData, // X-axis is just serial numbers
                    title: {
                        text: 'Serial No.'
                    }
                },

                legend: {
                    layout: 'vertical',
                    align: 'right',
                    verticalAlign: 'middle'
                },

                plotOptions: {
                    series: {
                        label: {
                            connectorAllowed: false
                        },
                        pointStart: 0
                    }
                },

                series: series,

                tooltip: {
                    shared: true, // This ensures multiple series are displayed
                    formatter: function () {
                        let tooltipText = '<b>Serial No. ' + this.x + '</b><br>';

                        // Loop through all the series and show their values
                        this.points.forEach(function (point) {
                            tooltipText += point.series.name + ': ' + point.y + '<br>';
                        });

                        return tooltipText;
                    }
                },

                responsive: {
                    rules: [{
                        condition: {
                            maxWidth: 500
                        },
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
        };

        reader.readAsArrayBuffer(file);
    }
});
