 class ExcelFilter {
            constructor() {
                this.workbook = null;
                this.worksheet = null;
                this.data = [];
                this.rawData = [];
                this.filteredData = [];
                this.isProcessing = false;
                this.headers = [];
                this.hiddenColumns = new Set();
                this.deletedRows = new Set();
                this.currentResults = [];
                this.columnWidths = {};
                this.correctHeaders = [
                    "счет", "сумма", "дата запуска", "поставщик", "-", "-", "дата оплат", "заказчик"
                ];
                this.debugMode = false;
                
                this.initializeElements();
                this.bindEvents();
            }

            initializeElements() {
                this.fileInput = document.getElementById('fileInput');
                this.searchInput = document.getElementById('searchInput');
                this.filterBtn = document.getElementById('filterBtn');
                this.resultsInfo = document.getElementById('resultsInfo');
                this.resultsContainer = document.getElementById('resultsContainer');
                this.fileInfo = document.getElementById('fileInfo');
                this.fileDetails = document.getElementById('fileDetails');
                this.progress = document.getElementById('progress');
                this.resetColsBtn = document.getElementById('resetColsBtn');
                this.exportBtn = document.getElementById('exportBtn');
                this.debugToggle = document.getElementById('debugToggle');
                this.debugInfo = document.getElementById('debugInfo');
                
                this.resizing = false;
                this.currentColumn = null;
                this.currentColIndex = null;
                this.startX = 0;
                this.startWidth = 0;
            }

            bindEvents() {
                this.fileInput.addEventListener('change', (e) => this.handleFileUpload(e));
                this.filterBtn.addEventListener('click', () => this.filterData());
                this.searchInput.addEventListener('keypress', (e) => {
                    if (e.key === 'Enter') {
                        this.filterData();
                    }
                });
                this.resetColsBtn.addEventListener('click', () => this.resetHiddenColumns());
                this.exportBtn.addEventListener('click', () => this.exportResults());
                this.debugToggle.addEventListener('click', () => this.toggleDebugMode());
                
                document.addEventListener('mousedown', this.handleMouseDown.bind(this));
                document.addEventListener('mousemove', this.handleMouseMove.bind(this));
                document.addEventListener('mouseup', this.handleMouseUp.bind(this));
            }

            toggleDebugMode() {
                this.debugMode = !this.debugMode;
                this.debugToggle.textContent = this.debugMode ? 'Отключить отладку' : 'Отладка';
                this.debugInfo.style.display = this.debugMode ? 'block' : 'none';
                if (this.debugMode) {
                    this.debugInfo.innerHTML = 'Режим отладки включен. Информация о поиске будет отображаться здесь.';
                }
            }

            async handleFileUpload(event) {
                const file = event.target.files[0];
                if (!file) return;

                this.showLoading('Загрузка файла...');
                
                try {
                    const data = await this.readFile(file);
                    this.workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: true,
                        cellText: false,
                        cellNF: true
                    });
                    
                    this.worksheet = this.workbook.Sheets[this.workbook.SheetNames[0]];
                    
                    this.rawData = XLSX.utils.sheet_to_json(this.worksheet, { 
                        header: 1,
                        defval: "",
                        blankrows: false,
                        raw: true
                    });
                    
                    this.processData();
                    
                    this.fileInfo.textContent = `Загружен файл: ${file.name} (${this.data.length} строк)`;
                    
                    this.filterBtn.disabled = false;
                    this.hideLoading();
                    
                } catch (error) {
                    this.showError('Ошибка при чтении файла: ' + error.message);
                    console.error('Error details:', error);
                }
            }

            processData() {
                this.data = this.rawData.filter(row => 
                    row.some(cell => cell !== null && cell !== "" && cell !== undefined && String(cell).trim() !== "")
                );
                
                if (this.data.length > 0) {
                    this.headers = this.data[0].map(header => 
                        header !== undefined && header !== null ? String(header) : '');
                    
                    for (let i = 8; i < this.headers.length; i++) {
                        this.headers[i] = "-";
                    }
                    
                    this.correctHeaders.forEach((correctHeader, index) => {
                        if (index < this.headers.length) {
                            this.headers[index] = correctHeader;
                        }
                    });
                    
                    this.data = this.data.slice(1);
                } else {
                    this.headers = [];
                }

                this.data = this.data.map(row => 
                    row.map((cell, index) => this.formatCellValue(cell, index))
                );
            }

            formatCellValue(value, columnIndex) {
                if (value === null || value === undefined || value === "") {
                    return "";
                }

                if (value instanceof Date) {
                    return this.formatDate(value);
                }

                if (typeof value === 'number') {
                    return this.formatNumber(value);
                }

                if (typeof value === 'string') {
                    const numberMatch = value.match(/^(\d+)\1$/);
                    if (numberMatch) {
                        return numberMatch[1];
                    }

                    const ipYearMatch = value.match(/^(ИП\s?\d{4})$/i);
                    if (ipYearMatch) {
                        return ipYearMatch[1];
                    }

                    const date = this.parseDateString(value);
                    if (date && !this.isSpecialFormat(value)) {
                        return this.formatDate(date);
                    }

                    const commaNumberMatch = value.match(/^(\d+),(\d+)$/);
                    if (commaNumberMatch) {
                        return `${commaNumberMatch[1]}.${commaNumberMatch[2]}`;
                    }

                    return value;
                }

                return value;
            }

            isSpecialFormat(value) {
                if (typeof value !== 'string') return false;
                return /^(ИП\s?\d{4})$/i.test(value);
            }

            formatDate(date) {
                const day = date.getDate().toString().padStart(2, '0');
                const month = (date.getMonth() + 1).toString().padStart(2, '0');
                const year = date.getFullYear();
                return `${day}.${month}.${year}`;
            }

            formatNumber(number) {
                if (Number.isInteger(number)) {
                    return number.toString();
                } else {
                    return Number(number.toFixed(2)).toString();
                }
            }

            parseDateString(value) {
                try {
                    if (value.includes('GMT') && value.length > 20) {
                        const datePart = value.split('GMT')[0].trim();
                        return new Date(datePart);
                    }
                    
                    const date = new Date(value);
                    if (!isNaN(date.getTime())) {
                        return date;
                    }
                } catch (e) {
                    return null;
                }
                return null;
            }

            readFile(file) {
                return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = (e) => resolve(e.target.result);
                    reader.onerror = (error) => reject(error);
                    reader.readAsArrayBuffer(file);
                });
            }

            async filterData() {
                if (this.isProcessing) return;
                
                const searchText = this.searchInput.value.trim();
                if (!searchText) {
                    this.showError('Введите текст для поиска');
                    return;
                }

                if (!this.data.length) {
                    this.showError('Сначала загрузите файл');
                    return;
                }

                this.isProcessing = true;
                this.filterBtn.disabled = true;
                this.showLoading('Поиск...');

                try {
                    this.filteredData = await this.processSearch(searchText);
                    this.displayResults();
                } catch (error) {
                    this.showError('Ошибка при фильтрации: ' + error.message);
                    console.error(error);
                } finally {
                    this.isProcessing = false;
                    this.filterBtn.disabled = false;
                    this.hideLoading();
                }
            }

            async processSearch(searchText) {
                const results = [];
                const total = this.data.length;
                const searchDigits = searchText.replace(/\D/g, '');
                const isPhoneSearch = searchDigits.length >= 7;
                const normalizedSearchText = searchText.replace(',', '.');
                
                let debugLog = [];
                
                for (let i = 0; i < total; i++) {
                    if (this.deletedRows.has(i)) continue;
                    
                    if (i % 100 === 0 || i === total - 1) {
                        this.progress.style.width = ((i / total) * 100) + '%';
                        await new Promise(resolve => setTimeout(resolve, 0));
                    }
                    
                    const row = this.data[i];
                    const isMatch = this.isRowMatch(row, searchText, normalizedSearchText, searchDigits, isPhoneSearch, debugLog);
                    
                    if (isMatch) {
                        results.push({row, originalIndex: i});
                    }
                }
                
                if (this.debugMode && debugLog.length > 0) {
                    this.debugInfo.innerHTML = '<strong>Информация отладки:</strong><br>' + debugLog.join('<br>');
                }
                
                this.progress.style.width = '100%';
                return results;
            }

            isRowMatch(row, searchText, normalizedSearchText, searchDigits, isPhoneSearch, debugLog = []) {
                for (let cell of row) {
                    if (cell !== null && cell !== undefined && cell !== "") {
                        const cellStr = String(cell);
                        const normalizedCell = cellStr.replace(',', '.');
                        
                        // Для числовых значений - более точное сравнение
                        if (this.isNumericValue(normalizedCell) && this.isNumericValue(normalizedSearchText)) {
                            const cellNum = parseFloat(normalizedCell);
                            const searchNum = parseFloat(normalizedSearchText);
                            
                            // Сравнение с учетом возможных ошибок округления
                            if (!isNaN(cellNum) && !isNaN(searchNum)) {
                                // Допустимая погрешность для чисел с плавающей точкой
                                const tolerance = 0.001;
                                const difference = Math.abs(cellNum - searchNum);
                                
                                if (difference < tolerance) {
                                    if (this.debugMode) {
                                        debugLog.push(`Найдено числовое совпадение: ${cellNum} ≈ ${searchNum} (разница: ${difference.toFixed(6)})`);
                                    }
                                    return true;
                                } else if (this.debugMode && searchText.includes('119334')) {
                                    debugLog.push(`Числовое сравнение: ${cellNum} ≠ ${searchNum} (разница: ${difference.toFixed(6)})`);
                                }
                            }
                        }
                        
                        if (isPhoneSearch) {
                            if (this.findPhoneInText(cellStr, searchDigits)) {
                                if (this.debugMode) {
                                    debugLog.push(`Найдено совпадение телефона: ${cellStr}`);
                                }
                                return true;
                            }
                        } else {
                            const cellLower = normalizedCell.toLowerCase();
                            const searchLower = searchText.toLowerCase();
                            const normalizedSearchLower = normalizedSearchText.toLowerCase();
                            
                            // Поиск по подстроке
                            if (cellLower.includes(searchLower) || 
                                cellLower.includes(normalizedSearchLower) ||
                                cellStr.includes(searchText) || 
                                cellStr.includes(normalizedSearchText)) {
                                if (this.debugMode) {
                                    debugLog.push(`Найдено текстовое совпадение: "${cellStr}" содержит "${searchText}"`);
                                }
                                return true;
                            }
                        }
                    }
                }
                return false;
            }

            isNumericValue(str) {
                return /^-?\d*\.?\d+$/.test(str);
            }

            findPhoneInText(text, searchDigits) {
                if (searchDigits.length < 7) return false;
                
                const phoneRegex = /[\d\+\(\)\-\s]{7,}/g;
                const matches = text.match(phoneRegex) || [];
                
                for (const match of matches) {
                    const matchDigits = match.replace(/\D/g, '');
                    
                    if (matchDigits.length >= 7) {
                        const normMatch = matchDigits.slice(-10);
                        const normSearch = searchDigits.slice(-10);
                        
                        if (normMatch === normSearch || 
                            normMatch.includes(normSearch) || 
                            normSearch.includes(normMatch)) {
                            return true;
                        }
                    }
                }
                
                return false;
            }

            displayResults() {
                this.resultsInfo.textContent = `Найдено совпадений: ${this.filteredData.length}`;
                
                if (this.filteredData.length === 0) {
                    this.resultsContainer.innerHTML = '<div class="error">Совпадений не найдено</div>';
                    return;
                }

                const searchText = this.searchInput.value.trim();
                const normalizedSearchText = searchText.replace(',', '.');
                const searchDigits = searchText.replace(/\D/g, '');
                const isPhoneSearch = searchDigits.length >= 7;
                
                let tableHTML = `
                    <div class="results-table-container">
                        <table class="results-table">
                            <thead>
                                <tr>
                `;
                
                this.headers.forEach((header, colIndex) => {
                    if (this.hiddenColumns.has(colIndex)) return;
                    
                    const headerClass = colIndex < this.correctHeaders.length ? 'corrected-header' : '';
                    const colWidth = this.columnWidths[colIndex] ? `style="width: ${this.columnWidths[colIndex]}px;"` : '';
                    
                    tableHTML += `
                        <th data-col="${colIndex}" ${colWidth}>
                            <div class="col-header">
                                <span class="${headerClass}">${this.escapeHtml(header)}</span>
                                <div class="col-actions">
                                    <button class="hide-col-btn" data-col="${colIndex}" title="Скрыть столбец">×</button>
                                </div>
                            </div>
                            <div class="col-resize-handle"></div>
                        </th>
                    `;
                });
                
                tableHTML += `
                                </tr>
                            </thead>
                            <tbody>
                `;
                
                this.currentResults = [];
                const displayData = this.filteredData.slice(0, 1000);
                
                displayData.forEach((item, rowIndex) => {
                    const row = item.row;
                    const originalIndex = item.originalIndex;
                    this.currentResults.push({row, originalIndex});
                    
                    tableHTML += '<tr data-row="' + rowIndex + '">';
                    this.headers.forEach((_, colIndex) => {
                        if (this.hiddenColumns.has(colIndex)) return;
                        
                        const value = row[colIndex] !== undefined ? row[colIndex] : '';
                        const cellWidth = this.columnWidths[colIndex] ? `style="width: ${this.columnWidths[colIndex]}px;"` : '';
                        const cellClass = this.isNumericValue(String(value).replace(',', '.')) ? 'number-cell' : '';
                        
                        tableHTML += `<td ${cellWidth} class="${cellClass}">${this.formatCellDisplay(value, searchText, normalizedSearchText, searchDigits, isPhoneSearch)}</td>`;
                    });
                    tableHTML += `
                        <td class="row-actions">
                            <button class="delete-row-btn" data-row="${rowIndex}" title="Удалить строку">×</button>
                        </td>
                    `;
                    tableHTML += '</tr>';
                });
                
                tableHTML += `
                            </tbody>
                        </table>
                    </div>
                `;
                
                if (this.filteredData.length > 1000) {
                    tableHTML += `<p class="success">Показано 1000 из ${this.filteredData.length} результатов</p>`;
                }
                
                this.resultsContainer.innerHTML = tableHTML;
                this.addTableEventListeners();
            }

            formatCellDisplay(value, searchText, normalizedSearchText, searchDigits, isPhoneSearch) {
                if (value === null || value === undefined || value === "") {
                    return "";
                }
                
                const text = String(value);
                
                if (!searchText) {
                    return this.escapeHtml(text);
                }
                
                if (isPhoneSearch) {
                    const phoneRegex = /[\d\+\(\)\-\s]{7,}/g;
                    let result = text;
                    
                    let match;
                    while ((match = phoneRegex.exec(text)) !== null) {
                        const phone = match[0];
                        const phoneDigits = phone.replace(/\D/g, '');
                        
                        if (phoneDigits.length >= 7) {
                            const normPhone = phoneDigits.slice(-10);
                            const normSearch = searchDigits.slice(-10);
                            
                            if (normPhone === normSearch || normPhone.includes(normSearch)) {
                                result = result.replace(
                                    phone, 
                                    `<span class="highlight">${this.escapeHtml(phone)}</span>`
                                );
                            }
                        }
                    }
                    
                    return result;
                } else {
                    const normalizedText = text.replace(',', '.');
                    const searchPatterns = [
                        searchText,
                        normalizedSearchText,
                        searchText.toLowerCase(),
                        normalizedSearchText.toLowerCase()
                    ];
                    
                    let result = text;
                    
                    for (const pattern of searchPatterns) {
                        if (pattern && normalizedText.includes(pattern)) {
                            const escapedPattern = this.escapeRegex(pattern);
                            const regex = new RegExp(escapedPattern, 'gi');
                            result = result.replace(regex, '<span class="highlight">$&</span>');
                        }
                    }
                    
                    return result;
                }
            }

            escapeRegex(string) {
                return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            }

            addTableEventListeners() {
                document.querySelectorAll('.hide-col-btn').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const colIndex = parseInt(e.target.dataset.col);
                        this.hideColumn(colIndex);
                        e.stopPropagation();
                    });
                });
                
                document.querySelectorAll('.delete-row-btn').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const rowIndex = parseInt(e.target.dataset.row);
                        this.deleteRow(rowIndex);
                        e.stopPropagation();
                    });
                });
            }
            
            hideColumn(colIndex) {
                this.hiddenColumns.add(colIndex);
                this.displayResults();
            }
            
            resetHiddenColumns() {
                this.hiddenColumns.clear();
                this.displayResults();
            }
            
            deleteRow(rowIndex) {
                if (rowIndex >= 0 && rowIndex < this.currentResults.length) {
                    const originalIndex = this.currentResults[rowIndex].originalIndex;
                    this.deletedRows.add(originalIndex);
                    this.filteredData = this.filteredData.filter((_, idx) => idx !== rowIndex);
                    this.displayResults();
                }
            }
            
            handleMouseDown(e) {
                if (e.target.classList.contains('col-resize-handle')) {
                    this.resizing = true;
                    this.currentColumn = e.target.parentElement;
                    this.currentColIndex = Array.from(this.currentColumn.parentElement.children).indexOf(this.currentColumn);
                    this.startX = e.clientX;
                    this.startWidth = this.currentColumn.offsetWidth;
                    
                    document.body.classList.add('resizing');
                    e.preventDefault();
                }
            }
            
            handleMouseMove(e) {
                if (!this.resizing || !this.currentColumn) return;
                
                const width = Math.max(30, this.startWidth + (e.clientX - this.startX));
                this.currentColumn.style.width = width + 'px';
                this.columnWidths[this.currentColIndex] = width;
                
                const cells = document.querySelectorAll(`.results-table td:nth-child(${this.currentColIndex + 1})`);
                cells.forEach(cell => {
                    cell.style.width = width + 'px';
                });
                
                e.preventDefault();
            }
            
            handleMouseUp() {
                if (this.resizing) {
                    this.resizing = false;
                    this.currentColumn = null;
                    this.currentColIndex = null;
                    document.body.classList.remove('resizing');
                }
            }
            
            exportResults() {
                if (this.filteredData.length === 0) {
                    this.showError('Нет данных для экспорта');
                    return;
                }
                
                try {
                    const wb = XLSX.utils.book_new();
                    const exportData = [];
                    
                    const visibleHeaders = this.headers.filter((_, index) => !this.hiddenColumns.has(index));
                    exportData.push(visibleHeaders);
                    
                    this.filteredData.forEach(item => {
                        const rowData = item.row.filter((_, index) => !this.hiddenColumns.has(index));
                        exportData.push(rowData);
                    });
                    
                    const ws = XLSX.utils.aoa_to_sheet(exportData);
                    XLSX.utils.book_append_sheet(wb, ws, "Результаты поиска");
                    XLSX.writeFile(wb, "результаты_поиска.xlsx");
                    
                } catch (error) {
                    this.showError('Ошибка при экспорте: ' + error.message);
                    console.error(error);
                }
            }

            escapeHtml(text) {
                if (text === null || text === undefined) return '';
                const div = document.createElement('div');
                div.textContent = text;
                return div.innerHTML;
            }

            showLoading(message) {
                this.resultsContainer.innerHTML = `<div class="loading">${message}</div>`;
            }

            hideLoading() {
                if (this.resultsContainer.querySelector('.loading')) {
                    this.resultsContainer.innerHTML = '';
                }
            }

            showError(message) {
                this.resultsContainer.innerHTML = `<div class="error">${message}</div>`;
            }
        }

        document.addEventListener('DOMContentLoaded', () => {
            new ExcelFilter();
        });