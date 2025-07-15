// 数据可视化工具主要JavaScript文件

class DataVisualizationTool {
    constructor() {
        this.workbook = null;
        this.currentData = null;
        this.charts = [];
        this.chartIdCounter = 0;
        
        this.initializeEventListeners();
    }

    // 初始化事件监听器
    initializeEventListeners() {
        // 文件选择
        document.getElementById('fileInput').addEventListener('change', this.handleFileSelect.bind(this));
        
        // 工作表选择
        document.getElementById('sheetSelect').addEventListener('change', this.handleSheetChange.bind(this));
        
        // 加载数据按钮
        document.getElementById('loadDataBtn').addEventListener('click', this.loadData.bind(this));
        
        // 图表类型变化
        document.getElementById('chartType').addEventListener('change', this.handleChartTypeChange.bind(this));
        
        // 生成图表按钮
        document.getElementById('generateChartBtn').addEventListener('click', this.generateChart.bind(this));
        
        // 布局类型变化
        document.getElementById('layoutType').addEventListener('change', this.updateChartsLayout.bind(this));
        
        // 清空所有图表
        document.getElementById('clearAllCharts').addEventListener('click', this.clearAllCharts.bind(this));
        
        // 帮助按钮
        document.getElementById('helpBtn').addEventListener('click', this.showHelp.bind(this));
    }

    // 处理文件选择
    async handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        try {
            this.showLoading(true);
            
            if (file.name.endsWith('.csv')) {
                await this.parseCSVFile(file);
            } else {
                await this.parseExcelFile(file);
            }
            
            this.showLoading(false);
        } catch (error) {
            this.showLoading(false);
            this.showError('文件解析失败: ' + error.message);
        }
    }

    // 解析CSV文件
    async parseCSVFile(file) {
        const text = await file.text();
        const lines = text.split('\n').filter(line => line.trim());
        
        const data = lines.map(line => {
            // 简单的CSV解析，处理引号内的逗号
            const result = [];
            let current = '';
            let inQuotes = false;
            
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                if (char === '"') {
                    inQuotes = !inQuotes;
                } else if (char === ',' && !inQuotes) {
                    result.push(current.trim());
                    current = '';
                } else {
                    current += char;
                }
            }
            result.push(current.trim());
            return result;
        });

        this.workbook = {
            SheetNames: ['Sheet1'],
            Sheets: {
                'Sheet1': this.arrayToSheet(data)
            }
        };

        this.setupSheetSelector(['Sheet1']);
        this.enableDataLoading();
    }

    // 解析Excel文件
    async parseExcelFile(file) {
        const arrayBuffer = await file.arrayBuffer();
        this.workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        this.setupSheetSelector(this.workbook.SheetNames);
        this.enableDataLoading();
    }

    // 数组转换为工作表格式
    arrayToSheet(data) {
        const ws = {};
        const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

        for (let R = 0; R < data.length; R++) {
            for (let C = 0; C < data[R].length; C++) {
                if (range.s.r > R) range.s.r = R;
                if (range.s.c > C) range.s.c = C;
                if (range.e.r < R) range.e.r = R;
                if (range.e.c < C) range.e.c = C;

                const cell = { v: data[R][C] };
                if (cell.v == null) continue;

                const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
                
                if (typeof cell.v === 'number') cell.t = 'n';
                else if (typeof cell.v === 'boolean') cell.t = 'b';
                else cell.t = 's';

                ws[cellRef] = cell;
            }
        }
        
        if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
        return ws;
    }

    // 设置工作表选择器
    setupSheetSelector(sheetNames) {
        const sheetSelect = document.getElementById('sheetSelect');
        const sheetSelector = document.getElementById('sheetSelector');
        
        sheetSelect.innerHTML = '';
        sheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        if (sheetNames.length > 1) {
            sheetSelector.style.display = 'block';
        } else {
            sheetSelector.style.display = 'none';
        }

        document.getElementById('dataRange').style.display = 'block';
    }

    // 启用数据加载
    enableDataLoading() {
        document.getElementById('loadDataBtn').disabled = false;
    }

    // 处理工作表变化
    handleSheetChange() {
        // 工作表变化时可以在这里添加预处理逻辑
    }

    // 加载数据
    loadData() {
        try {
            const sheetName = document.getElementById('sheetSelect').value;
            const startRow = parseInt(document.getElementById('startRow').value) || 1;
            const endRow = parseInt(document.getElementById('endRow').value) || 100;
            const hasHeader = document.getElementById('hasHeader').checked;

            const worksheet = this.workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                range: `A${startRow}:ZZ${endRow}`
            });

            if (jsonData.length === 0) {
                this.showError('没有找到数据');
                return;
            }

            this.currentData = this.processData(jsonData, hasHeader);
            this.displayDataPreview();
            this.setupChartConfiguration();
            
        } catch (error) {
            this.showError('数据加载失败: ' + error.message);
        }
    }

    // 处理数据
    processData(rawData, hasHeader) {
        const data = {
            headers: [],
            rows: [],
            raw: rawData
        };

        if (hasHeader && rawData.length > 0) {
            data.headers = rawData[0].map((header, index) => 
                header || `列${index + 1}`
            );
            data.rows = rawData.slice(1);
        } else {
            data.headers = rawData[0] ? rawData[0].map((_, index) => `列${index + 1}`) : [];
            data.rows = rawData;
        }

        // 过滤空行
        data.rows = data.rows.filter(row => 
            row && row.some(cell => cell !== null && cell !== undefined && cell !== '')
        );

        return data;
    }

    // 显示数据预览
    displayDataPreview() {
        const dataPreview = document.getElementById('dataPreview');
        const dataTable = document.getElementById('dataTable');
        const dataInfo = document.getElementById('dataInfo');

        // 更新数据信息
        dataInfo.textContent = `共 ${this.currentData.rows.length} 行数据，${this.currentData.headers.length} 列`;

        // 创建表格
        let tableHTML = '<thead><tr>';
        this.currentData.headers.forEach(header => {
            tableHTML += `<th>${header}</th>`;
        });
        tableHTML += '</tr></thead><tbody>';

        // 只显示前20行数据
        const displayRows = this.currentData.rows.slice(0, 20);
        displayRows.forEach(row => {
            tableHTML += '<tr>';
            this.currentData.headers.forEach((_, index) => {
                const cellValue = row[index] || '';
                tableHTML += `<td>${cellValue}</td>`;
            });
            tableHTML += '</tr>';
        });

        if (this.currentData.rows.length > 20) {
            tableHTML += `<tr><td colspan="${this.currentData.headers.length}" class="text-center text-muted">
                ... 还有 ${this.currentData.rows.length - 20} 行数据未显示
            </td></tr>`;
        }

        tableHTML += '</tbody>';
        dataTable.innerHTML = tableHTML;
        dataPreview.style.display = 'block';
        dataPreview.classList.add('fade-in');
    }

    // 设置图表配置
    setupChartConfiguration() {
        const xAxisColumn = document.getElementById('xAxisColumn');
        const yAxisColumns = document.getElementById('yAxisColumns');

        // 清空选项
        xAxisColumn.innerHTML = '';
        yAxisColumns.innerHTML = '';

        // 添加列选项
        this.currentData.headers.forEach((header, index) => {
            const xOption = document.createElement('option');
            xOption.value = index;
            xOption.textContent = header;
            xAxisColumn.appendChild(xOption);

            const yOption = document.createElement('option');
            yOption.value = index;
            yOption.textContent = header;
            yAxisColumns.appendChild(yOption);
        });

        // 默认选择
        if (this.currentData.headers.length > 0) {
            xAxisColumn.value = 0;
        }
        if (this.currentData.headers.length > 1) {
            yAxisColumns.value = 1;
            yAxisColumns.options[1].selected = true;
        }

        document.getElementById('chartConfig').style.display = 'block';
        document.getElementById('chartConfig').classList.add('fade-in');
    }

    // 处理图表类型变化
    handleChartTypeChange() {
        const chartType = document.getElementById('chartType').value;
        const xAxisColumn = document.getElementById('xAxisColumn').parentElement;
        const yAxisColumns = document.getElementById('yAxisColumns').parentElement;

        // 根据图表类型调整UI
        if (chartType === 'pie' || chartType === 'treemap' || chartType === 'funnel') {
            xAxisColumn.style.display = 'block';
            yAxisColumns.style.display = 'block';
            document.getElementById('yAxisColumns').multiple = false;
        } else if (chartType === 'gauge') {
            xAxisColumn.style.display = 'none';
            yAxisColumns.style.display = 'block';
            document.getElementById('yAxisColumns').multiple = false;
        } else {
            xAxisColumn.style.display = 'block';
            yAxisColumns.style.display = 'block';
            document.getElementById('yAxisColumns').multiple = true;
        }
    }

    // 生成图表
    generateChart() {
        try {
            const chartType = document.getElementById('chartType').value;
            const xAxisColumnIndex = parseInt(document.getElementById('xAxisColumn').value);
            const yAxisColumnsIndices = Array.from(document.getElementById('yAxisColumns').selectedOptions)
                .map(option => parseInt(option.value));
            const chartTitle = document.getElementById('chartTitle').value || 
                `${this.getChartTypeName(chartType)} - ${new Date().toLocaleTimeString()}`;

            if (yAxisColumnsIndices.length === 0) {
                this.showError('请选择Y轴数据列');
                return;
            }

            const chartData = this.prepareChartData(chartType, xAxisColumnIndex, yAxisColumnsIndices);
            const chartConfig = this.createChartConfig(chartType, chartData, chartTitle);
            
            this.createChartContainer(chartConfig, chartTitle, chartType);
            this.updateChartManager();

        } catch (error) {
            this.showError('图表生成失败: ' + error.message);
        }
    }

    // 准备图表数据
    prepareChartData(chartType, xAxisColumnIndex, yAxisColumnsIndices) {
        const xAxisData = [];
        const seriesData = [];

        // 初始化系列数据
        yAxisColumnsIndices.forEach(columnIndex => {
            seriesData.push({
                name: this.currentData.headers[columnIndex],
                data: [],
                columnIndex: columnIndex
            });
        });

        // 处理数据行
        this.currentData.rows.forEach(row => {
            const xValue = row[xAxisColumnIndex];
            if (xValue === null || xValue === undefined || xValue === '') return;

            xAxisData.push(xValue);

            yAxisColumnsIndices.forEach((columnIndex, seriesIndex) => {
                const yValue = this.parseNumber(row[columnIndex]);
                seriesData[seriesIndex].data.push(yValue);
            });
        });

        return {
            xAxisData,
            seriesData,
            xAxisName: this.currentData.headers[xAxisColumnIndex]
        };
    }

    // 解析数字
    parseNumber(value) {
        if (value === null || value === undefined || value === '') return 0;
        const num = parseFloat(value);
        return isNaN(num) ? 0 : num;
    }

    // 创建图表配置
    createChartConfig(chartType, chartData, title) {
        const baseConfig = {
            title: {
                text: title,
                left: 'center',
                textStyle: {
                    fontSize: 16,
                    fontWeight: 'bold'
                }
            },
            tooltip: {
                trigger: 'axis',
                axisPointer: {
                    type: 'shadow'
                }
            },
            legend: {
                top: '10%',
                left: 'center'
            },
            grid: {
                left: '3%',
                right: '4%',
                bottom: '3%',
                containLabel: true
            }
        };

        switch (chartType) {
            case 'line':
                return this.createLineChartConfig(baseConfig, chartData);
            case 'bar':
                return this.createBarChartConfig(baseConfig, chartData);
            case 'pie':
                return this.createPieChartConfig(baseConfig, chartData);
            case 'scatter':
                return this.createScatterChartConfig(baseConfig, chartData);
            case 'radar':
                return this.createRadarChartConfig(baseConfig, chartData);
            case 'heatmap':
                return this.createHeatmapChartConfig(baseConfig, chartData);
            case 'treemap':
                return this.createTreemapChartConfig(baseConfig, chartData);
            case 'funnel':
                return this.createFunnelChartConfig(baseConfig, chartData);
            case 'gauge':
                return this.createGaugeChartConfig(baseConfig, chartData);
            default:
                return this.createLineChartConfig(baseConfig, chartData);
        }
    }

    // 创建折线图配置
    createLineChartConfig(baseConfig, chartData) {
        return {
            ...baseConfig,
            xAxis: {
                type: 'category',
                data: chartData.xAxisData,
                name: chartData.xAxisName
            },
            yAxis: {
                type: 'value'
            },
            series: chartData.seriesData.map(series => ({
                name: series.name,
                type: 'line',
                data: series.data,
                smooth: true,
                symbolSize: 8
            }))
        };
    }

    // 创建柱状图配置
    createBarChartConfig(baseConfig, chartData) {
        return {
            ...baseConfig,
            xAxis: {
                type: 'category',
                data: chartData.xAxisData,
                name: chartData.xAxisName
            },
            yAxis: {
                type: 'value'
            },
            series: chartData.seriesData.map(series => ({
                name: series.name,
                type: 'bar',
                data: series.data,
                itemStyle: {
                    borderRadius: [4, 4, 0, 0]
                }
            }))
        };
    }

    // 创建饼图配置
    createPieChartConfig(baseConfig, chartData) {
        const pieData = chartData.xAxisData.map((name, index) => ({
            name: name,
            value: chartData.seriesData[0].data[index]
        }));

        return {
            ...baseConfig,
            tooltip: {
                trigger: 'item',
                formatter: '{a} <br/>{b}: {c} ({d}%)'
            },
            series: [{
                name: chartData.seriesData[0].name,
                type: 'pie',
                radius: ['40%', '70%'],
                data: pieData,
                emphasis: {
                    itemStyle: {
                        shadowBlur: 10,
                        shadowOffsetX: 0,
                        shadowColor: 'rgba(0, 0, 0, 0.5)'
                    }
                }
            }]
        };
    }

    // 创建散点图配置
    createScatterChartConfig(baseConfig, chartData) {
        const scatterData = chartData.seriesData.length >= 2 ? 
            chartData.seriesData[0].data.map((x, index) => [x, chartData.seriesData[1].data[index]]) :
            chartData.seriesData[0].data.map((y, index) => [index, y]);

        return {
            ...baseConfig,
            xAxis: {
                type: 'value',
                name: chartData.seriesData[0].name
            },
            yAxis: {
                type: 'value',
                name: chartData.seriesData[1] ? chartData.seriesData[1].name : 'Y轴'
            },
            series: [{
                name: '散点数据',
                type: 'scatter',
                data: scatterData,
                symbolSize: 10
            }]
        };
    }

    // 创建雷达图配置
    createRadarChartConfig(baseConfig, chartData) {
        const indicator = chartData.xAxisData.map(name => ({ name, max: Math.max(...chartData.seriesData[0].data) * 1.2 }));
        
        return {
            ...baseConfig,
            radar: {
                indicator: indicator,
                center: ['50%', '60%'],
                radius: '70%'
            },
            series: chartData.seriesData.map(series => ({
                name: series.name,
                type: 'radar',
                data: [{
                    value: series.data,
                    name: series.name
                }],
                areaStyle: {
                    opacity: 0.3
                }
            }))
        };
    }

    // 创建热力图配置
    createHeatmapChartConfig(baseConfig, chartData) {
        const heatmapData = [];
        chartData.xAxisData.forEach((xValue, xIndex) => {
            chartData.seriesData.forEach((series, yIndex) => {
                heatmapData.push([xIndex, yIndex, series.data[xIndex] || 0]);
            });
        });

        return {
            ...baseConfig,
            xAxis: {
                type: 'category',
                data: chartData.xAxisData
            },
            yAxis: {
                type: 'category',
                data: chartData.seriesData.map(s => s.name)
            },
            visualMap: {
                min: 0,
                max: Math.max(...heatmapData.map(d => d[2])),
                calculable: true,
                orient: 'horizontal',
                left: 'center',
                bottom: '10%'
            },
            series: [{
                name: '热力图',
                type: 'heatmap',
                data: heatmapData,
                label: {
                    show: true
                }
            }]
        };
    }

    // 创建树图配置
    createTreemapChartConfig(baseConfig, chartData) {
        const treeData = chartData.xAxisData.map((name, index) => ({
            name: name,
            value: chartData.seriesData[0].data[index]
        }));

        return {
            ...baseConfig,
            series: [{
                name: '树图',
                type: 'treemap',
                data: treeData,
                leafDepth: 1,
                label: {
                    show: true,
                    formatter: '{b}: {c}'
                },
                itemStyle: {
                    borderColor: '#fff'
                }
            }]
        };
    }

    // 创建漏斗图配置
    createFunnelChartConfig(baseConfig, chartData) {
        const funnelData = chartData.xAxisData.map((name, index) => ({
            name: name,
            value: chartData.seriesData[0].data[index]
        })).sort((a, b) => b.value - a.value);

        return {
            ...baseConfig,
            series: [{
                name: '漏斗图',
                type: 'funnel',
                left: '10%',
                top: 60,
                bottom: 60,
                width: '80%',
                min: 0,
                max: Math.max(...funnelData.map(d => d.value)),
                minSize: '0%',
                maxSize: '100%',
                sort: 'descending',
                gap: 2,
                label: {
                    show: true,
                    position: 'inside'
                },
                data: funnelData
            }]
        };
    }

    // 创建仪表盘配置
    createGaugeChartConfig(baseConfig, chartData) {
        const value = chartData.seriesData[0].data[0] || 0;
        const max = Math.max(...chartData.seriesData[0].data) * 1.2;

        return {
            ...baseConfig,
            series: [{
                name: '仪表盘',
                type: 'gauge',
                detail: {
                    formatter: '{value}',
                    fontSize: 20
                },
                data: [{
                    value: value,
                    name: chartData.seriesData[0].name
                }],
                max: max,
                axisLine: {
                    lineStyle: {
                        width: 20,
                        color: [
                            [0.3, '#67e0e3'],
                            [0.7, '#37a2da'],
                            [1, '#fd666d']
                        ]
                    }
                }
            }]
        };
    }

    // 创建图表容器
    createChartContainer(chartConfig, title, type) {
        const chartId = `chart_${this.chartIdCounter++}`;
        const chartsContainer = document.getElementById('chartsContainer');
        const emptyState = document.getElementById('emptyState');

        // 隐藏空状态
        if (emptyState) {
            emptyState.style.display = 'none';
        }

        // 创建图表容器HTML
        const chartHTML = `
            <div class="chart-container fade-in" id="${chartId}_container">
                <div class="chart-header">
                    <h6 class="chart-title">${title}</h6>
                    <div class="chart-controls">
                        <button class="btn btn-sm btn-outline-light" onclick="app.downloadChart('${chartId}')">
                            <i class="bi bi-download"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-light" onclick="app.removeChart('${chartId}')">
                            <i class="bi bi-x"></i>
                        </button>
                    </div>
                </div>
                <div class="chart-content">
                    <div id="${chartId}" class="chart-canvas"></div>
                </div>
            </div>
        `;

        chartsContainer.insertAdjacentHTML('beforeend', chartHTML);

        // 初始化ECharts图表
        const chartElement = document.getElementById(chartId);
        const chart = echarts.init(chartElement);
        chart.setOption(chartConfig);

        // 保存图表信息
        this.charts.push({
            id: chartId,
            instance: chart,
            config: chartConfig,
            title: title,
            type: type
        });

        // 更新布局
        this.updateChartsLayout();

        // 显示图表管理面板
        document.getElementById('chartManager').style.display = 'block';

        // 响应式处理
        window.addEventListener('resize', () => {
            chart.resize();
        });
    }

    // 更新图表布局
    updateChartsLayout() {
        const layoutType = document.getElementById('layoutType').value;
        const chartsContainer = document.getElementById('chartsContainer');
        
        // 移除所有布局类
        chartsContainer.classList.remove('charts-grid', 'charts-stack', 'charts-horizontal');
        
        // 添加新的布局类
        if (this.charts.length > 0) {
            chartsContainer.classList.add(`charts-${layoutType}`);
        }

        // 调整图表尺寸
        setTimeout(() => {
            this.charts.forEach(chart => {
                chart.instance.resize();
            });
        }, 100);
    }

    // 更新图表管理器
    updateChartManager() {
        const chartList = document.getElementById('chartList');
        chartList.innerHTML = '';

        this.charts.forEach(chart => {
            const chartItem = document.createElement('div');
            chartItem.className = 'chart-item';
            chartItem.innerHTML = `
                <div>
                    <div class="chart-item-title">${chart.title}</div>
                    <div class="chart-item-type">${this.getChartTypeName(chart.type)}</div>
                </div>
                <div class="chart-item-controls">
                    <button class="btn btn-sm btn-outline-primary" onclick="app.downloadChart('${chart.id}')">
                        <i class="bi bi-download"></i>
                    </button>
                    <button class="btn btn-sm btn-outline-danger" onclick="app.removeChart('${chart.id}')">
                        <i class="bi bi-trash"></i>
                    </button>
                </div>
            `;
            chartList.appendChild(chartItem);
        });
    }

    // 获取图表类型名称
    getChartTypeName(type) {
        const typeNames = {
            'line': '折线图',
            'bar': '柱状图',
            'pie': '饼图',
            'scatter': '散点图',
            'radar': '雷达图',
            'heatmap': '热力图',
            'treemap': '树图',
            'funnel': '漏斗图',
            'gauge': '仪表盘'
        };
        return typeNames[type] || type;
    }

    // 下载图表
    downloadChart(chartId) {
        const chart = this.charts.find(c => c.id === chartId);
        if (chart) {
            const url = chart.instance.getDataURL({
                type: 'png',
                pixelRatio: 2,
                backgroundColor: '#fff'
            });
            
            const link = document.createElement('a');
            link.download = `${chart.title}.png`;
            link.href = url;
            link.click();
        }
    }

    // 移除图表
    removeChart(chartId) {
        const chartIndex = this.charts.findIndex(c => c.id === chartId);
        if (chartIndex !== -1) {
            // 销毁ECharts实例
            this.charts[chartIndex].instance.dispose();
            
            // 移除DOM元素
            const chartContainer = document.getElementById(`${chartId}_container`);
            if (chartContainer) {
                chartContainer.remove();
            }
            
            // 从数组中移除
            this.charts.splice(chartIndex, 1);
            
            // 更新图表管理器
            this.updateChartManager();
            
            // 如果没有图表了，显示空状态
            if (this.charts.length === 0) {
                document.getElementById('emptyState').style.display = 'block';
                document.getElementById('chartManager').style.display = 'none';
            }
        }
    }

    // 清空所有图表
    clearAllCharts() {
        if (this.charts.length === 0) return;
        
        if (confirm('确定要清空所有图表吗？')) {
            this.charts.forEach(chart => {
                chart.instance.dispose();
                const chartContainer = document.getElementById(`${chart.id}_container`);
                if (chartContainer) {
                    chartContainer.remove();
                }
            });
            
            this.charts = [];
            this.updateChartManager();
            document.getElementById('emptyState').style.display = 'block';
            document.getElementById('chartManager').style.display = 'none';
        }
    }

    // 显示帮助
    showHelp() {
        const helpModal = new bootstrap.Modal(document.getElementById('helpModal'));
        helpModal.show();
    }

    // 显示加载状态
    showLoading(show) {
        const loadBtn = document.getElementById('loadDataBtn');
        if (show) {
            loadBtn.disabled = true;
            loadBtn.innerHTML = '<i class="bi bi-arrow-repeat"></i> 加载中...';
        } else {
            loadBtn.disabled = false;
            loadBtn.innerHTML = '<i class="bi bi-arrow-down-circle"></i> 加载数据';
        }
    }

    // 显示错误信息
    showError(message) {
        alert('错误: ' + message);
    }
}

// 初始化应用
const app = new DataVisualizationTool();

// 确保在窗口大小改变时调整图表
window.addEventListener('resize', () => {
    app.charts.forEach(chart => {
        chart.instance.resize();
    });
});