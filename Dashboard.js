<script>
  // ApexCharts instances
  let admissionsTrendChart;
  let admissionsByRlcRegionChart;
  let genderDistributionChart;
  let admissionsByQualificationChart;
  let lowCompletionCentersChart;
  let examEventPredictionChart;
  let alcLlcBubbleChart;
  let admissionsByAgeGroupChart;
  let admissionsTrendPerRegionGenderChart;
  let admissionsHeatmapChart; // New heatmap chart instance

  // Call loadDashboard when the page is loaded
  document.addEventListener('DOMContentLoaded', loadDashboard);

  function loadDashboard() {
    console.log("Loading dashboard...");
    // Initialize current year filter to current year
    document.getElementById('filterYear').value = new Date().getFullYear();

    // Fetch unique filter values and then the dashboard data
    google.script.run
      .withSuccessHandler(populateFiltersAndInitialData)
      .withFailureHandler(showError)
      .getUniqueFilterValues();
  }

  function populateSelect(id, values, defaultValue = 'All') {
    const select = document.getElementById(id);
    select.innerHTML = ''; // Clear existing options

    const allOption = document.createElement('option');
    allOption.value = 'All';
    allOption.textContent = `All ${id.replace('filter', '').replace(/([A-Z])/g, ' $1').trim()}`;
    select.appendChild(allOption);

    values.forEach(value => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });

    if (defaultValue) {
        select.value = defaultValue;
    }
    // Special handling for filterYear to set to current year by default
    if (id === 'filterYear') {
        const currentYear = new Date().getFullYear().toString();
        if (values.includes(currentYear)) {
            select.value = currentYear;
        } else if (values.length > 0) {
             select.value = values[values.length -1]; // Select latest available year
        } else {
             select.value = 'All';
        }
    }
  }

  function populateFiltersAndInitialData(filterValues) {
    console.log("Populating filters...", filterValues);
    populateSelect('filterYear', filterValues.years); // Default to latest year handled in populateSelect
    populateSelect('filterExamEvent', filterValues.examEvents);
    populateSelect('filterRlcRegion', filterValues.rlcRegions);
    populateSelect('filterLearnerDistrict', filterValues.learnerDistricts);
    populateSelect('filterGender', filterValues.genders);
    populateSelect('filterQualification', filterValues.qualifications);

    // Set a default value for the selected exam event for prediction, e.g., the latest one or a specific one
    if (filterValues.examEvents && filterValues.examEvents.length > 0) {
        // Ensure the main filterExamEvent is set to the latest if available
        document.getElementById('filterExamEvent').value = filterValues.examEvents[filterValues.examEvents.length -1];
    }
    // Set a default for the exam event end date, e.g., end of current year
    const today = new Date();
    document.getElementById('examEventEndDate').value = new Date(today.getFullYear(), 11, 31).toISOString().split('T')[0];

    // Now fetch the actual dashboard data with initial filters
    applyFilters();
  }

  function applyFilters() {
    const filters = {
      year: document.getElementById('filterYear').value,
      examEvent: document.getElementById('filterExamEvent').value,
      rlcRegion: document.getElementById('filterRlcRegion').value,
      learnerDistrict: document.getElementById('filterLearnerDistrict').value,
      gender: document.getElementById('filterGender').value,
      qualification: document.getElementById('filterQualification').value,
      overallTarget: document.getElementById('overallTarget').value,
      manualTarget: document.getElementById('manualTarget').value,
      selectedExamEvent: document.getElementById('filterExamEvent').value, // Use the main exam event filter for prediction, assuming consistency
      examEventEndDate: document.getElementById('examEventEndDate').value,
      yoyTarget: document.getElementById('yoyTarget').value
    };

    console.log("Applying filters:", filters);
    google.script.run
      .withSuccessHandler(updateDashboard)
      .withFailureHandler(showError)
      .getDataForDashboard(filters);
  }

  function updateDashboard(data) {
    console.log("Dashboard data received:", data);
    updateKPIs(data.kpis);
    renderCharts(data.chartData); // overallTarget is now passed via KPI updates or directly in chart data if needed
    updateCenterPerformanceTable(data.tableData);
    displayRecommendations(data.recommendations);
  }

  function updateKPIs(kpis) {
    document.getElementById('kpiTotalAdmissions').textContent = kpis.totalAdmissions.toLocaleString();
    document.getElementById('kpiCompletionRate').textContent = kpis.completionRate;
    document.getElementById('kpiPredictedAdmissions').textContent = kpis.predictedAdmissions !== 'N/A' ? kpis.predictedAdmissions.toLocaleString() : kpis.predictedAdmissions;

    const predictionWarningElement = document.getElementById('kpiPredictionWarning');
    const predictedAdmissionsCard = document.querySelector('.kpi-card.bg-danger'); // Select the card for predicted admissions

    if (kpis.onTrackWarning) {
      predictionWarningElement.innerHTML = `<small class="text-white-75">Predicted to be below target of ${kpis.overallTarget.toLocaleString()}.</small>`;
      predictedAdmissionsCard.style.background = 'linear-gradient(135deg, #dc3545, #b02a37)'; // Red for below target
    } else if (kpis.predictedAdmissions !== 'N/A' && kpis.predictedAdmissions >= kpis.overallTarget * 0.9) {
      predictionWarningElement.innerHTML = `<small class="text-white-75">On track for target of ${kpis.overallTarget.toLocaleString()}.</small>`;
      predictedAdmissionsCard.style.background = 'linear-gradient(135deg, #28a745, #218838)'; // Green for on track
    } else {
      predictionWarningElement.textContent = ''; // Clear warning if N/A or default (already handled by first if)
      predictedAdmissionsCard.style.background = 'linear-gradient(135deg, #dc3545, #b02a37)'; // Default red if no specific warning
    }

    // Update YOY Growth KPI
    document.getElementById('kpiYoyGrowth').textContent = kpis.yoyGrowth !== 'N/A' ? `${kpis.yoyGrowth}%` : kpis.yoyGrowth;
    document.getElementById('kpiYoyTarget').textContent = `${kpis.yoyTarget}%`;
    const yoyGrowthStatus = document.getElementById('kpiYoyGrowthStatus');
    const yoyGrowthCard = document.querySelector('.kpi-card.bg-primary'); // Select the card for YOY growth

    if (kpis.yoyGrowth !== 'N/A') {
        const growth = parseFloat(kpis.yoyGrowth);
        const target = parseFloat(kpis.yoyTarget);
        if (growth >= target) {
            yoyGrowthStatus.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" class="bi bi-arrow-up-circle-fill" viewBox="0 0 16 16"><path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0m-7.5 3.5a.5.5 0 0 0 1 0V5.707l2.146 2.147a.5.5 0 0 0 .708-.708l-3-3a.5.5 0 0 0-.708 0l-3 3a.5.5 0 1 0 .708.708L7.5 5.707z"/></svg>';
            yoyGrowthCard.style.background = 'linear-gradient(135deg, #28a745, #218838)'; // Green for good growth
        } else {
            yoyGrowthStatus.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" class="bi bi-arrow-down-circle-fill" viewBox="0 0 16 16"><path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0m-7.5-3.5a.5.5 0 0 0-1 0v5.793L5.354 8.646a.5.5 0 1 0-.708.708l3 3a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 11.293z"/></svg>';
            yoyGrowthCard.style.background = 'linear-gradient(135deg, #dc3545, #b02a37)'; // Red for below target growth
        }
    } else {
        yoyGrowthStatus.innerHTML = ''; // Clear status icon if no data
        yoyGrowthCard.style.background = 'linear-gradient(135deg, #007bff, #0056b3)'; // Default blue
    }
  }

  function renderCharts(chartData) { // Removed overallTarget from params as it's now handled in KPI/chart data
    // Chart 1: Admissions Trend by Month
    const admissionsTrendOptions = {
      series: chartData.admissionsTrend.series,
      chart: {
        height: '100%', // Use 100% height for parent
        type: 'line',
        zoom: { enabled: false },
        toolbar: { show: false }
      },
      dataLabels: { enabled: false },
      stroke: { curve: 'smooth' },
      title: { text: 'Monthly Admissions', align: 'left' },
      grid: { row: { colors: ['#f3f3f3', 'transparent'], opacity: 0.5 } },
      xaxis: { categories: chartData.admissionsTrend.categories },
      tooltip: { enabled: true }
    };
    if (admissionsTrendChart) { admissionsTrendChart.updateOptions(admissionsTrendOptions); }
    else { admissionsTrendChart = new ApexCharts(document.querySelector("#admissionsTrendChart"), admissionsTrendOptions); admissionsTrendChart.render(); }

    // Chart 2: Admissions by RLC Region
    const admissionsByRlcRegionOptions = {
      series: [{ data: chartData.admissionsByRlcRegion }],
      chart: {
        height: '100%', // Use 100% height for parent
        type: 'bar',
        toolbar: { show: false }
      },
      plotOptions: { bar: { horizontal: false, columnWidth: '55%', endingShape: 'rounded' } },
      dataLabels: { enabled: false },
      stroke: { show: true, width: 2, colors: ['transparent'] },
      xaxis: { categories: chartData.admissionsByRlcRegion.map(item => item.x) },
      yaxis: { title: { text: 'Admissions' } },
      fill: { opacity: 1 },
      tooltip: {
        y: {
          formatter: function (val) { return val + " admissions" }
        }
      }
    };
    if (admissionsByRlcRegionChart) { admissionsByRlcRegionChart.updateOptions(admissionsByRlcRegionOptions); }
    else { admissionsByRlcRegionChart = new ApexCharts(document.querySelector("#admissionsByRlcRegionChart"), admissionsByRlcRegionOptions); admissionsByRlcRegionChart.render(); }

    // Chart 3: Gender Distribution (Pie/Donut Chart)
    const genderDistributionOptions = {
        series: chartData.genderDistribution.map(item => item.y),
        chart: {
            width: '100%', // Use 100% width
            height: '100%', // Use 100% height
            type: 'donut',
        },
        labels: chartData.genderDistribution.map(item => item.x),
        responsive: [{
            breakpoint: 480,
            options: {
                chart: {
                    width: '100%'
                },
                legend: {
                    position: 'bottom'
                }
            }
        }],
        legend: {
            position: 'bottom' // Always put legend at bottom for better space usage
        }
    };
    if (genderDistributionChart) { genderDistributionChart.updateOptions(genderDistributionOptions); }
    else { genderDistributionChart = new ApexCharts(document.querySelector("#genderDistributionChart"), genderDistributionOptions); genderDistributionChart.render(); }

    // Chart 4: Admissions by Qualification (Bar Chart)
    const admissionsByQualificationOptions = {
      series: [{ data: chartData.admissionsByQualification }],
      chart: {
        height: '100%', // Use 100% height for parent
        type: 'bar',
        toolbar: { show: false }
      },
      plotOptions: { bar: { horizontal: false, columnWidth: '55%', endingShape: 'rounded' } },
      dataLabels: { enabled: false },
      stroke: { show: true, width: 2, colors: ['transparent'] },
      xaxis: { categories: chartData.admissionsByQualification.map(item => item.x) },
      yaxis: { title: { text: 'Admissions' } },
      fill: { opacity: 1 },
      tooltip: {
        y: {
          formatter: function (val) { return val + " admissions" }
        }
      }
    };
    if (admissionsByQualificationChart) { admissionsByQualificationChart.updateOptions(admissionsByQualificationOptions); }
    else { admissionsByQualificationChart = new ApexCharts(document.querySelector("#admissionsByQualificationChart"), admissionsByQualificationOptions); admissionsByQualificationChart.render(); }

    // Chart 5: Low-Completion Centers (Horizontal Bar Chart)
    const lowCompletionCentersOptions = {
      series: [{ name: 'Completion %', data: chartData.lowCompletionCenters }],
      chart: {
        height: 250, // Fixed smaller height for this specific chart
        type: 'bar',
        toolbar: { show: false }
      },
      plotOptions: { bar: { horizontal: true } },
      dataLabels: { enabled: true, formatter: (val) => `${val}%` },
      xaxis: {
        categories: chartData.lowCompletionCenters.map(item => item.x),
        title: { text: 'Completion %' },
        max: 100 // Completion percentage
      },
      yaxis: { title: { text: 'Center' } },
      title: { text: 'Top 10 Underperforming Centers (<90% Completion)', align: 'left' }
    };
    if (lowCompletionCentersChart) { lowCompletionCentersChart.updateOptions(lowCompletionCentersOptions); }
    else { lowCompletionCentersChart = new ApexCharts(document.querySelector("#lowCompletionCentersChart"), lowCompletionCentersOptions); lowCompletionCentersChart.render(); }


    // Chart 6: Exam Event Target Prediction
    const examEventPredictionOptions = {
      series: chartData.examEventPrediction.series,
      chart: {
        type: 'bar',
        height: '100%', // Use 100% height for parent
        toolbar: { show: false }
      },
      plotOptions: {
        bar: {
          horizontal: false,
          columnWidth: '50%',
          endingShape: 'rounded'
        },
      },
      dataLabels: { enabled: false },
      stroke: { show: true, width: 2, colors: ['transparent'] },
      xaxis: { categories: ['Admissions'] },
      yaxis: { title: { text: 'Number of Admissions' } },
      fill: { opacity: 1 },
      tooltip: {
        y: {
          formatter: function (val) { return val + " admissions" }
        }
      },
      annotations: {
        yaxis: [{
          y: chartData.examEventPrediction.displayedTarget,
          borderColor: '#00E396',
          label: {
            borderColor: '#00E396',
            style: {
              color: '#fff',
              background: '#00E396',
            },
            text: `Target: ${chartData.examEventPrediction.displayedTarget.toLocaleString()}`,
          }
        }]
      }
    };
    if (examEventPredictionChart) { examEventPredictionChart.updateOptions(examEventPredictionOptions); }
    else { examEventPredictionChart = new ApexCharts(document.querySelector("#examEventPredictionChart"), examEventPredictionOptions); examEventPredictionChart.render(); }
    document.getElementById('examEventTargetAlert').textContent = chartData.examEventPrediction.targetAlert;


    // Chart 7: Geographic & Demographic Admissions Heatmap
    const admissionsHeatmapOptions = {
        series: chartData.admissionsHeatmap.series,
        chart: {
            height: '100%', // Use 100% height for parent
            type: 'heatmap',
            toolbar: { show: false }
        },
        dataLabels: {
            enabled: false
        },
        colors: ["#008FFB"], // You can customize heatmap colors
        xaxis: {
            type: 'category',
            categories: chartData.admissionsHeatmap.demoCategories, // Demographic categories on X-axis
            labels: {
                formatter: function (val) {
                    return val.length > 10 ? val.substring(0, 10) + '...' : val; // Truncate long labels
                }
            }
        },
        yaxis: {
            categories: chartData.admissionsHeatmap.geoCategories, // Geographic categories on Y-axis
            labels: {
                formatter: function (val) {
                    return val.length > 15 ? val.substring(0, 15) + '...' : val; // Truncate long labels
                }
            }
        },
        title: {
            text: 'Admissions Heatmap (Region vs. Qualification)',
            align: 'left'
        },
        grid: { padding: { right: 20 } },
        tooltip: {
            x: { formatter: (val) => val },
            y: {
                formatter: function(val, { series, seriesIndex, dataPointIndex, w }) {
                    const geo = w.globals.initialSeries[seriesIndex].name;
                    const demo = w.globals.labels[dataPointIndex];
                    return `Region: ${geo}, Qualification: ${demo}, Admissions: ${val}`;
                }
            }
        }
    };
    if (admissionsHeatmapChart) { admissionsHeatmapChart.updateOptions(admissionsHeatmapOptions); }
    else { admissionsHeatmapChart = new ApexCharts(document.querySelector("#admissionsHeatmap"), admissionsHeatmapOptions); admissionsHeatmapChart.render(); }


  // ... inside the renderCharts function

// ... inside the renderCharts function

    // START: REPLACED SECTION FOR CHART 8
    // Chart 8: ALC/LLC Performance as a Scatter Plot with Quadrants
    // This is a much better visualization than a bubble chart for this data.
    const alcLlcScatterPlotOptions = {
        // We can reuse the same data structure
        series: chartData.alcLlcBubbleChart, 
        chart: {
            height: '100%',
            type: 'scatter', // CHANGED to scatter
            zoom: {
                enabled: true,
                type: 'xy'
            },
            toolbar: {
                show: true
            }
        },
        legend: {
            show: false // Legend is not useful here
        },
        title: {
            text: 'Center Performance Matrix: Completion Rate vs. Average Age',
            align: 'left'
        },
        subtitle: {
            text: 'Hover on a point for center details. Lines show median age & 90% completion.'
        },
        xaxis: {
            type: 'numeric',
            title: {
                text: 'Average Learner Age'
            },
            labels: {
                formatter: function (val) {
                    return val.toFixed(0);
                }
            },
            // Let the chart determine the min/max for the best fit
        },
        yaxis: {
            min: 0,
            max: 100,
            tickAmount: 5,
            title: {
                text: 'Completion Rate (%)'
            },
             labels: {
                formatter: function (val) {
                    return val.toFixed(0) + '%';
                }
            }
        },
        // Add annotations to create performance quadrants
        annotations: {
            yaxis: [{
                y: 90, // Target completion rate
                borderColor: '#dc3545',
                strokeDashArray: 4,
                label: {
                    borderColor: '#dc3545',
                    style: {
                        color: '#fff',
                        background: '#dc3545'
                    },
                    text: '90% Target'
                }
            }],
            // Note: A vertical line for median age would require calculating the median in Code.gs
            // For simplicity, we'll stick to the completion target line which is highly effective.
        },
        tooltip: {
            enabled: true,
            custom: function({ series, seriesIndex, dataPointIndex, w }) {
                // The custom tooltip from the previous attempt is perfect here
                const data = w.globals.initialSeries[seriesIndex].data[dataPointIndex];
                const centerName = w.globals.initialSeries[seriesIndex].name;
                const avgAge = data[0];
                const completion = data[1];
                const admissions = data[2]; // We can still get this from the data!

                return `<div class="p-2">
                            <div class="fw-bold mb-1">${centerName}</div>
                            <hr class="my-1">
                            <div>Completion Rate: <span class="fw-bold">${completion}%</span></div>
                            <div>Average Age: <span class="fw-bold">${avgAge}</span></div>
                            <div>Total Admissions: <span class="fw-bold">${admissions}</span></div>
                        </div>`;
            }
        }
    };
    // Make sure to update the chart instance logic
    if (alcLlcBubbleChart) { alcLlcBubbleChart.updateOptions(alcLlcScatterPlotOptions); }
    else { alcLlcBubbleChart = new ApexCharts(document.querySelector("#alcLlcBubbleChart"), alcLlcScatterPlotOptions); alcLlcBubbleChart.render(); }
    // END: REPLACED SECTION FOR CHART 8




    // Chart 9: Admissions by Age Group (Bar Chart)
    const admissionsByAgeGroupOptions = {
        series: [{ data: chartData.admissionsByAgeGroup }],
        chart: {
            height: '100%', // Use 100% height for parent
            type: 'bar',
            toolbar: { show: false }
        },
        plotOptions: { bar: { horizontal: false, columnWidth: '55%', endingShape: 'rounded' } },
        dataLabels: { enabled: false },
        xaxis: { categories: chartData.admissionsByAgeGroup.map(item => item.x) },
        yaxis: { title: { text: 'Admissions' } },
        fill: { opacity: 1 },
        title: { text: 'Admissions by Age Group', align: 'left' }
    };
    if (admissionsByAgeGroupChart) { admissionsByAgeGroupChart.updateOptions(admissionsByAgeGroupOptions); }
    else { admissionsByAgeGroupChart = new ApexCharts(document.querySelector("#admissionsByAgeGroupChart"), admissionsByAgeGroupOptions); admissionsByAgeGroupChart.render(); }

    // Chart 10: Admissions Trend Per Region (Gender Split - Stacked Bar Chart)
    const admissionsTrendPerRegionGenderOptions = {
        series: chartData.admissionsTrendPerRegionGender.series,
        chart: {
            height: '100%', // Use 100% height for parent
            type: 'bar',
            stacked: true,
            toolbar: { show: false }
        },
        plotOptions: {
            bar: {
                horizontal: false,
                dataLabels: {
                    total: {
                        enabled: true,
                        offsetX: 0,
                        style: {
                            fontSize: '13px',
                            fontWeight: 900
                        }
                    }
                }
            },
        },
        stroke: {
            width: 1,
            colors: ['#fff']
        },
        title: {
            text: 'Admissions by RLC Region and Gender',
            align: 'left'
        },
        xaxis: {
            categories: chartData.admissionsTrendPerRegionGender.categories,
            title: { text: 'RLC Region' } // Changed X-axis title
        },
        yaxis: {
            title: { text: 'Number of Admissions' }
        },
        tooltip: {
            y: {
                formatter: function (val) {
                    return val + " admissions"
                }
            }
        },
        fill: {
            opacity: 1
        },
        legend: {
            position: 'top',
            horizontalAlign: 'left',
            offsetX: 40
        }
    };
    if (admissionsTrendPerRegionGenderChart) { admissionsTrendPerRegionGenderChart.updateOptions(admissionsTrendPerRegionGenderOptions); }
    else { admissionsTrendPerRegionGenderChart = new ApexCharts(document.querySelector("#admissionsTrendPerRegionGenderChart"), admissionsTrendPerRegionGenderOptions); admissionsTrendPerRegionGenderChart.render(); }
  }


 function updateCenterPerformanceTable(tableData) {
    const tbody = document.querySelector('#centerPerformanceTable tbody');
    tbody.innerHTML = ''; // Clear existing rows

    // Filter for underperforming centers (e.g., completion < 90%) and sort them
    const underperformingCenters = tableData
        .filter(row => parseFloat(row.completionPct) < 90)
        .sort((a, b) => a.completionPct - b.completionPct); // Sort by lowest completion first

    if (underperformingCenters.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" class="text-center">No underperforming centers found for the selected filters.</td></tr>';
      return;
    }

    underperformingCenters.forEach(row => {
      const tr = document.createElement('tr');
      // Add a visual cue for very low completion rates if desired
      const completionClass = parseFloat(row.completionPct) < 70 ? 'text-danger fw-bold' : '';
      tr.innerHTML = `
        <td>${row.centerName}</td>
        <td>${row.actualAdmissions.toLocaleString()}</td>
        <td class="${completionClass}">${row.completionPct}%</td>
        <td>${row.avgInternalMarks}</td>
        <td>${row.avgInternalScore}</td>
        <td>${row.genderMix}</td>
        <td>${row.paymentStatusMix}</td>
      `;
      tbody.appendChild(tr);
    });
}

  function displayRecommendations(recommendations) {
    const ul = document.getElementById('recommendationsList');
    ul.innerHTML = ''; // Clear existing recommendations

    if (recommendations.length === 0) {
      const li = document.createElement('li');
      li.className = 'list-group-item';
      li.textContent = 'No specific recommendations at this time. Dashboard looks healthy!';
      ul.appendChild(li);
      return;
    }

    recommendations.forEach(rec => {
      const li = document.createElement('li');
      li.className = 'list-group-item';
      li.textContent = rec;
      ul.appendChild(li);
    });
  }

  function showError(error) {
    console.error("Error from Apps Script:", error);
    alert("An error occurred: " + error.message || error);
  }
</script>