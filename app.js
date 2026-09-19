// Real demand curves over 24 hours (l/s) matching Excel average magnitudes
        const PATTERN_SALIDA_1_ALCALA = [18.5, 17.5, 16.0, 15.5, 15.0, 16.5, 21.0, 26.0, 31.0, 33.5, 32.5, 31.5, 30.5, 29.8, 29.5, 30.2, 31.8, 33.2, 34.5, 32.8, 29.8, 26.2, 22.5, 19.5];
        const PATTERN_BURGUILLOS = [12.0, 11.5, 10.5, 10.5, 10.5, 11.5, 12.8, 14.5, 15.8, 16.8, 16.5, 15.2, 14.6, 14.0, 13.8, 14.4, 15.2, 16.1, 16.8, 15.8, 14.6, 13.7, 12.8, 12.2];
        const PATTERN_SALIDA_1_ENTRONQUE = [22.0, 20.5, 19.0, 18.5, 18.0, 19.5, 24.0, 29.0, 34.0, 36.5, 35.5, 34.5, 33.5, 32.8, 32.5, 33.2, 34.8, 36.2, 37.5, 35.8, 32.8, 29.2, 25.5, 22.5];
        const PATTERN_ENTRADA_ALCALA_DEFAULT = [40.0, 40.0, 40.0, 40.0, 40.0, 40.0, 42.0, 46.0, 49.0, 49.0, 49.0, 46.0, 43.0, 43.0, 43.0, 45.0, 47.0, 49.0, 49.0, 46.0, 43.0, 43.0, 40.0, 40.0];
        const PATTERN_ENTRADA_ENTRONQUE_DEFAULT = [45.0, 45.0, 45.0, 45.0, 45.0, 45.0, 47.0, 51.0, 54.0, 54.0, 54.0, 51.0, 48.0, 48.0, 48.0, 50.0, 52.0, 54.0, 54.0, 51.0, 48.0, 48.0, 45.0, 45.0];

        // Global Chart References
        let chartAlcNivel = null;
        let chartAlcCaudales = null;
        let chartEntNivel = null;
        let chartEntCaudales = null;

        // Modal State
        let activeModalTarget = 'alcala';
        let tempHourlyProfile = {
            entrada: [],
            salida1: [],
            burguillos: []
        };

        // Global State
        const state = {
            currentTab: 'alcala',
            alcala: {
                burguillosOption: 'con',
                useDefaultInlet: true,
                nivelInicio: 6.00,
                nivelMinimo: 1.00,
                caudalEntradaMedio: 44.36,
                isCustomData: false,
                startDate: null,
                timestamps: [],
                entradaProfile: [...PATTERN_ENTRADA_ALCALA_DEFAULT],
                salida1Profile: [...PATTERN_SALIDA_1_ALCALA],
                burguillosProfile: [...PATTERN_BURGUILLOS],
                factor: 1329,
                maxHours: 168,
                visibleHours: 168,
                breachHour: -1,
                simulation: []
            },
            entronque: {
                useDefaultInlet: true,
                nivelInicio: 5.50,
                nivelMinimo: 1.50,
                caudalEntradaMedio: 48.50,
                isCustomData: false,
                startDate: null,
                timestamps: [],
                entradaProfile: [...PATTERN_ENTRADA_ENTRONQUE_DEFAULT],
                salida1Profile: [...PATTERN_SALIDA_1_ENTRONQUE],
                factor: 1798,
                maxHours: 168,
                visibleHours: 168,
                breachHour: -1,
                simulation: []
            }
        };

        function normalizeExcelDate(value) {
            let date = null;
            if (value instanceof Date && !Number.isNaN(value.getTime())) {
                date = new Date(value.getTime());
            } else if (typeof value === 'number') {
                const parsed = XLSX.SSF.parse_date_code(value);
                if (parsed) date = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, Math.floor(parsed.S || 0), 0);
            } else if (typeof value === 'string' && value.trim()) {
                const direct = new Date(value.trim());
                if (!Number.isNaN(direct.getTime())) date = direct;
            }
            if (!date || Number.isNaN(date.getTime())) return null;
            date.setMilliseconds(0);
            return date;
        }

        function formatDateTime(value) {
            const date = normalizeExcelDate(value);
            if (!date) return '';
            const pad = value => String(value).padStart(2, '0');
            return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
        }

        function getTimeLabels(target, hours) {
            const base = state[target].startDate;
            if (base instanceof Date && !Number.isNaN(base.getTime())) {
                return Array.from({ length: hours }, (_, i) => formatDateTime(new Date(base.getTime() + i * 3600000)));
            }
            return Array.from({ length: hours }, (_, i) => `Día ${Math.floor(i / 24) + 1} ${String(i % 24).padStart(2, '0')}:00`);
        }

        function formatAutonomy(hours) {
            if (hours === null || hours === undefined || hours >= 168) return '> 7 días';
            const days = Math.floor(hours / 24);
            const remainingHours = Math.floor(hours % 24);
            const parts = [];
            if (days) parts.push(`${days} ${days === 1 ? 'día' : 'días'}`);
            if (remainingHours) parts.push(`${remainingHours} h`);
            return parts.length ? parts.join(' ') : '0 h';
        }

        function getInletProfileAverage(target) {
            const targetState = state[target];
            const profile = Array.isArray(targetState.entradaProfile) ? targetState.entradaProfile : [];
            const periodHours = Math.min(168, profile.length);
            if (!periodHours) return Number(targetState.caudalEntradaMedio) || 0;
            const total = profile.slice(0, periodHours).reduce((sum, value) => sum + Number(value || 0), 0);
            return total / periodHours;
        }

        function syncManeuverDateInput(target) {
            const id = target === 'alcala' ? 'alc-fecha-inicio' : 'ent-fecha-inicio';
            const input = document.getElementById(id);
            const date = state[target].startDate;
            if (!input || !(date instanceof Date) || Number.isNaN(date.getTime())) return;
            const pad = value => String(value).padStart(2, '0');
            input.value = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
        }


        function syncInletAverageField(target) {
            const inputId = target === 'alcala' ? 'alc-caudal-entrada' : 'ent-caudal-entrada';
            const input = document.getElementById(inputId);
            if (!input || !state[target].useDefaultInlet) return;
            const average = getInletProfileAverage(target);
            state[target].caudalEntradaMedio = average;
            input.value = average.toFixed(2);
        }

        function setManeuverStart(target){const id=target==='alcala'?'alc-fecha-inicio':'ent-fecha-inicio',v=document.getElementById(id).value,d=v?new Date(v):null;state[target].startDate=d&&!Number.isNaN(d.getTime())?d:null;target==='alcala'?updateAlcalaSimulation():updateEntronqueSimulation();}
        async function printTab(target) {
            document.body.classList.remove('print-alcala', 'print-entronque');
            document.body.classList.add(`print-${target}`);
            const previousTitle = document.title;
            document.title = `Informe hidráulico - ${target === 'alcala' ? 'Alcalá del Río' : 'Entronque'}`;

            if (target === 'alcala') {
                updateAlcalaCharts();
            } else {
                updateEntronqueCharts();
            }

            // ApexCharts necesita terminar el render específico de impresión.
            await new Promise(resolve => setTimeout(resolve, 900));
            window.print();

            setTimeout(() => {
                document.title = previousTitle;
                document.body.classList.remove(`print-${target}`);
                if (target === 'alcala') updateAlcalaCharts(); else updateEntronqueCharts();
            }, 900);
        }

        // Switch between Excel Profile (Checked) and Constant Manual Input (Unchecked)
        function toggleInletMode(target) {
            if (target === 'alcala') {
                const isChecked = document.getElementById('alc-check-default-entrada').checked;
                state.alcala.useDefaultInlet = isChecked;
                const inputEl = document.getElementById('alc-caudal-entrada');

                if (isChecked) {
                    inputEl.disabled = true;
                    inputEl.className = 'w-full bg-slate-100 border border-slate-300 rounded-lg px-3 py-2 text-sm font-semibold text-slate-500 cursor-not-allowed focus:ring-2 focus:ring-brand-500 focus:outline-none';
                    const avg = getInletProfileAverage('alcala').toFixed(2);
                    inputEl.value = avg;
                } else {
                    inputEl.disabled = false;
                    inputEl.className = 'w-full bg-white border border-slate-300 rounded-lg px-3 py-2 text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-brand-500 focus:outline-none';
                    inputEl.focus();
                }
                updateAlcalaSimulation();
            } else {
                const isChecked = document.getElementById('ent-check-default-entrada').checked;
                state.entronque.useDefaultInlet = isChecked;
                const inputEl = document.getElementById('ent-caudal-entrada');

                if (isChecked) {
                    inputEl.disabled = true;
                    inputEl.className = 'w-full bg-slate-100 border border-slate-300 rounded-lg px-3 py-2 text-sm font-semibold text-slate-500 cursor-not-allowed focus:ring-2 focus:ring-indigo-500 focus:outline-none';
                    const avg = getInletProfileAverage('entronque').toFixed(2);
                    inputEl.value = avg;
                } else {
                    inputEl.disabled = false;
                    inputEl.className = 'w-full bg-white border border-slate-300 rounded-lg px-3 py-2 text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none';
                    inputEl.focus();
                }
                updateEntronqueSimulation();
            }
        }

        // Hydraulic calculation engine for Alcalá del Río
        function calculateAlcala() {
            const data = [];
            let currentNivel = parseFloat(state.alcala.nivelInicio) || 0;
            const factor = state.alcala.factor;
            const isBurguillosActive = state.alcala.burguillosOption === 'con';
            const useExcelProfile = state.alcala.useDefaultInlet;
            const fixedInletVal = parseFloat(document.getElementById('alc-caudal-entrada').value) || 0;
            const minNivel = parseFloat(state.alcala.nivelMinimo) || 0;
            const maxHours = state.alcala.startDate ? Math.min(168, state.alcala.entradaProfile.length) : state.alcala.maxHours;
            const labels = getTimeLabels('alcala', maxHours);
            let breachHour = -1;

            for (let i = 0; i < maxHours; i++) {
                const hourOfDay = i % 24;
                const dayNum = Math.floor(i / 24) + 1;
                const entrada = useExcelProfile ? (state.alcala.entradaProfile[i] ?? state.alcala.entradaProfile[hourOfDay] ?? 44.36) : fixedInletVal;
                const salida1 = state.alcala.salida1Profile[i] ?? state.alcala.salida1Profile[hourOfDay] ?? 29.9;
                const burguillos = isBurguillosActive ? (state.alcala.burguillosProfile[i] ?? state.alcala.burguillosProfile[hourOfDay] ?? 0) : 0;
                const salidaTotal = salida1 + burguillos;
                const netFlow = entrada - salidaTotal;
                const nextNivel = ((netFlow * 3.6) + (currentNivel * factor)) / factor;

                if (breachHour === -1 && nextNivel <= minNivel) breachHour = i + 1;
                data.push({ hora: labels[i], dayNum, hourOfDay, entrada, salida1, burguillos, salidaTotal, netFlow, nivel: nextNivel });
                currentNivel = nextNivel;
            }

            let visibleHours = maxHours;
            if (breachHour !== -1) visibleHours = Math.min(maxHours, Math.max(24, breachHour + 3));
            state.alcala.visibleHours = visibleHours;
            state.alcala.breachHour = breachHour;
            state.alcala.simulation = data.slice(0, visibleHours);
        }

        // Hydraulic calculation engine for Entronque
        function calculateEntronque() {
            const data = [];
            let currentNivel = parseFloat(state.entronque.nivelInicio) || 0;
            const factor = state.entronque.factor;
            const useExcelProfile = state.entronque.useDefaultInlet;
            const fixedInletVal = parseFloat(document.getElementById('ent-caudal-entrada').value) || 0;
            const minNivel = parseFloat(state.entronque.nivelMinimo) || 0;
            const maxHours = state.entronque.startDate ? Math.min(168, state.entronque.entradaProfile.length) : state.entronque.maxHours;
            const labels = getTimeLabels('entronque', maxHours);
            let breachHour = -1;

            for (let i = 0; i < maxHours; i++) {
                const hourOfDay = i % 24;
                const dayNum = Math.floor(i / 24) + 1;
                const entrada = useExcelProfile ? (state.entronque.entradaProfile[i] ?? state.entronque.entradaProfile[hourOfDay] ?? 48.5) : fixedInletVal;
                const salida1 = state.entronque.salida1Profile[i] ?? state.entronque.salida1Profile[hourOfDay] ?? 30.0;
                const netFlow = entrada - salida1;
                const nextNivel = ((netFlow * 3.6) + (currentNivel * factor)) / factor;

                if (breachHour === -1 && nextNivel <= minNivel) breachHour = i + 1;
                data.push({ hora: labels[i], dayNum, hourOfDay, entrada, salida1, burguillos: 0, salidaTotal: salida1, netFlow, nivel: nextNivel });
                currentNivel = nextNivel;
            }

            let visibleHours = maxHours;
            if (breachHour !== -1) visibleHours = Math.min(maxHours, Math.max(24, breachHour + 3));
            state.entronque.visibleHours = visibleHours;
            state.entronque.breachHour = breachHour;
            state.entronque.simulation = data.slice(0, visibleHours);
        }

        function calculateAlcalaScenario(withBurguillos){let level=parseFloat(state.alcala.nivelInicio)||0;const initial=level,alertLevel=parseFloat(state.alcala.nivelMinimo)||0,factor=state.alcala.factor,hours=state.alcala.startDate?Math.min(168,state.alcala.entradaProfile.length):state.alcala.maxHours,fixed=parseFloat(document.getElementById('alc-caudal-entrada').value)||0,labels=getTimeLabels('alcala',hours);let breach=-1;const ins=[],outs=[];for(let i=0;i<hours;i++){const hod=i%24,input=state.alcala.useDefaultInlet?(state.alcala.entradaProfile[i]??state.alcala.entradaProfile[hod]??44.36):fixed,out1=state.alcala.salida1Profile[i]??state.alcala.salida1Profile[hod]??29.9,b=withBurguillos?(state.alcala.burguillosProfile[i]??state.alcala.burguillosProfile[hod]??0):0;ins.push(input);outs.push(out1+b);level=(((input-out1-b)*3.6)+(level*factor))/factor;if(breach<0&&level<=alertLevel)breach=i;}const n=breach<0?hours:breach+1,avg=v=>v.slice(0,n).reduce((x,y)=>x+y,0)/n;return{initial,alertLevel,inletAvg:breach<0?getInletProfileAverage('alcala'):avg(ins),outletAvg:avg(outs),autonomy:breach<0?'> 7 días':formatAutonomy(n),minimumTime:breach<0?'No alcanzado':labels[breach]};}
        function updateAlcalaHypothesisSummary(){for(const[k,v]of Object.entries({con:calculateAlcalaScenario(true),sin:calculateAlcalaScenario(false)})){const put=(f,t)=>{const el=document.getElementById(`hyp-${k}-${f}`);if(el)el.innerText=t;};put('inicio',`${v.initial.toFixed(2)} m`);put('alerta',`${v.alertLevel.toFixed(2)} m`);put('entrada',`${v.inletAvg.toFixed(2)} l/s`);put('salida',`${v.outletAvg.toFixed(2)} l/s`);put('hora-minimo',v.minimumTime);put('autonomia',v.autonomy);}}

        // Update UI for Alcalá
        function updateAlcalaUI(){calculateAlcala();updateAlcalaCharts();updateAlcalaHypothesisSummary();}
        function isPrintMode() { return document.body.classList.contains('print-alcala') || document.body.classList.contains('print-entronque'); }

        function smoothFlowSeries(values, radius = 8) {
            if (!Array.isArray(values) || values.length === 0) return [];
            const smoothPass = input => input.map((value, index) => {
                let weightedSum = 0;
                let weightTotal = 0;
                for (let offset = -radius; offset <= radius; offset++) {
                    const sourceIndex = Math.min(input.length - 1, Math.max(0, index + offset));
                    const weight = radius + 1 - Math.abs(offset);
                    weightedSum += input[sourceIndex] * weight;
                    weightTotal += weight;
                }
                return weightedSum / weightTotal;
            });
            return smoothPass(smoothPass(values)).map(value => Number(value.toFixed(2)));
        }

        function updateAlcalaCharts() {
            const sim = state.alcala.simulation;
            if (!sim || sim.length === 0) return;

            const labels = sim.map(s => s.hora);
            const niveles = sim.map(s => Number(s.nivel.toFixed(3)));
            const minLimit = parseFloat(state.alcala.nivelMinimo);
            const maxValue = Math.max(...niveles);
            const minValue = Math.min(...niveles);
            const maxIndex = niveles.indexOf(maxValue);
            const minIndex = niveles.indexOf(minValue);

            const levelOptions = {
                chart: { type: 'line', height: isPrintMode() ? 390 : 420, width: '100%', parentHeightOffset: 0, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, selection: { enabled: false }, dropShadow: { enabled: false }, animations: { enabled: false }, zoom: { enabled: false }, selection: { enabled: false }, dropShadow: { enabled: false }, sparkline: { enabled: false } },
                series: [{ name: 'Nivel del depósito', data: niveles }, { name: 'Nivel mínimo de alerta', data: Array(labels.length).fill(minLimit) }],
                colors: ['#0284c7','#ef4444'],
                stroke: { curve: 'smooth', width: [3,2], dashArray: [0,6] },
                markers: { size: 0, hover: { size: 7, sizeOffset: 3 } },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: isPrintMode() ? 6 : 10, tickPlacement: 'between', labels: { rotate: isPrintMode() ? 0 : -90, rotateAlways: !isPrintMode(), hideOverlappingLabels: true, trim: false, formatter: value => isPrintMode() && value ? String(value).replace(' ', '\n') : value, style: { fontSize: isPrintMode() ? '6px' : '10px', colors: '#64748b' }, offsetY: isPrintMode() ? 4 : 0 }, tooltip: { enabled: false } },
                yaxis: { min: Math.max(0, Math.floor(Math.min(minValue, minLimit) - 0.5)), max: Math.ceil(maxValue + 0.5), labels: { formatter: v => `${v.toFixed(2)} m`, style: { colors: '#64748b' } } },
                tooltip: { shared: false, intersect: false, followCursor: true, x: { show: true }, y: { formatter: v => `${v.toFixed(2)} m` } },
                legend: {
                    show: true,
                    position: 'bottom',
                    horizontalAlign: 'center',
                    floating: false,
                    fontSize: isPrintMode() ? '8px' : '12px',
                    fontWeight: 600,
                    offsetX: 0,
                    offsetY: isPrintMode() ? 10 : 0,
                    itemMargin: {
                        horizontal: 10,
                        vertical: 0
                    }
                },
                grid: {
                    borderColor: '#e2e8f0',
                    padding: {
                        top: 5,
                        right: 14,
                        bottom: isPrintMode() ? 24 : 0,
                        left: 14
                    }
                },
                annotations: { yaxis: [] }
            };

            const flowOptions = {
                chart: { type: 'line', height: isPrintMode() ? 390 : 420, width: '100%', parentHeightOffset: 0, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, selection: { enabled: false }, dropShadow: { enabled: false }, animations: { enabled: false }, zoom: { enabled: false }, selection: { enabled: false }, dropShadow: { enabled: false }, sparkline: { enabled: false } },
                series: [
                    { name: 'Entrada', data: smoothFlowSeries(sim.map(s => Number(s.entrada)), 8) },
                    { name: 'Salida', data: smoothFlowSeries(sim.map(s => Number(s.salidaTotal)), 8) }
                ],
                colors: ['#10b981', '#e11d48'],
                stroke: { curve: 'smooth', width: [3, 3], lineCap: 'round' },
                markers: {
                    size: 0,
                    strokeWidth: 2,
                    strokeColors: '#ffffff',
                    hover: {
                        size: 7,
                        sizeOffset: 3
                    }
                },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: isPrintMode() ? 6 : 10, tickPlacement: 'between', labels: { rotate: isPrintMode() ? 0 : -90, rotateAlways: !isPrintMode(), hideOverlappingLabels: true, trim: false, formatter: value => isPrintMode() && value ? String(value).replace(' ', '\n') : value, style: { fontSize: isPrintMode() ? '6px' : '10px', colors: '#64748b' }, offsetY: isPrintMode() ? 4 : 0 }, tooltip: { enabled: false } },
                yaxis: { labels: { formatter: v => v.toFixed(1), style: { colors: '#64748b' } } },
                tooltip: {
                    enabled: true,
                    shared: true,
                    intersect: false,
                    followCursor: true,
                    hideEmptySeries: true,
                    marker: {
                        show: true
                    },
                    x: {
                        show: true
                    },
                    y: {
                        formatter: value => {
                            if (value === undefined || value === null) {
                                return '';
                            }
                             return `${value.toFixed(2)} l/s`;
                        }
                    }
                },
                legend: {
                    show: true,
                    position: 'bottom',
                    horizontalAlign: 'center',
                    floating: false,
                    fontSize: isPrintMode() ? '8px' : '12px',
                    fontWeight: 600,
                    offsetX: 0,
                    offsetY: isPrintMode() ? 10 : 0,
                    itemMargin: {
                        horizontal: 10,
                        vertical: 0
                    }
                },
                grid: {
                    borderColor: '#e2e8f0',
                    padding: {
                        top: 5,
                        right: 14,
                        bottom: isPrintMode() ? 24 : 0,
                        left: 14
                    }
                },
            };

            if (chartAlcNivel) chartAlcNivel.destroy();
            chartAlcNivel = new ApexCharts(document.querySelector('#chart-alcala-nivel'), levelOptions);
            chartAlcNivel.render();

            if (chartAlcCaudales) chartAlcCaudales.destroy();
            chartAlcCaudales = new ApexCharts(document.querySelector('#chart-alcala-caudales'), flowOptions);
            chartAlcCaudales.render();
        }

        // Update UI for Entronque
        function updateEntronqueUI(){calculateEntronque();updateEntronqueCharts();}
        function updateEntronqueCharts() {
            const sim = state.entronque.simulation;
            if (!sim || sim.length === 0) return;

            const labels = sim.map(s => s.hora);
            const niveles = sim.map(s => Number(s.nivel.toFixed(3)));
            const minLimit = parseFloat(state.entronque.nivelMinimo);
            const maxValue = Math.max(...niveles);
            const minValue = Math.min(...niveles);
            const maxIndex = niveles.indexOf(maxValue);
            const minIndex = niveles.indexOf(minValue);

            const levelOptions = {
                chart: { type: 'line', height: isPrintMode() ? 390 : 420, width: '100%', parentHeightOffset: 0, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, selection: { enabled: false }, dropShadow: { enabled: false }, animations: { enabled: false }, zoom: { enabled: false }, selection: { enabled: false }, dropShadow: { enabled: false }, sparkline: { enabled: false } },
                series: [{ name: 'Nivel del depósito', data: niveles }, { name: 'Nivel mínimo de alerta', data: Array(labels.length).fill(minLimit) }],
                colors: ['#6366f1','#ef4444'],
                stroke: { curve: 'smooth', width: [3,2], dashArray: [0,6] },
                markers: { size: 0, hover: { size: 7, sizeOffset: 3 } },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: isPrintMode() ? 6 : 10, tickPlacement: 'between', labels: { rotate: isPrintMode() ? 0 : -90, rotateAlways: !isPrintMode(), hideOverlappingLabels: true, trim: false, formatter: value => isPrintMode() && value ? String(value).replace(' ', '\n') : value, style: { fontSize: isPrintMode() ? '6px' : '10px', colors: '#64748b' }, offsetY: isPrintMode() ? 4 : 0 }, tooltip: { enabled: false } },
                yaxis: { min: Math.max(0, Math.floor(Math.min(minValue, minLimit) - 0.5)), max: Math.ceil(maxValue + 0.5), labels: { formatter: v => `${v.toFixed(2)} m`, style: { colors: '#64748b' } } },
                tooltip: { shared: false, intersect: false, followCursor: true, x: { show: true }, y: { formatter: v => `${v.toFixed(2)} m` } },
                legend: {
                    show: true,
                    position: 'bottom',
                    horizontalAlign: 'center',
                    floating: false,
                    fontSize: isPrintMode() ? '8px' : '12px',
                    fontWeight: 600,
                    offsetX: 0,
                    offsetY: isPrintMode() ? 10 : 0,
                    itemMargin: {
                        horizontal: 10,
                        vertical: 0
                    }
                },
                grid: {
                    borderColor: '#e2e8f0',
                    padding: {
                        top: 5,
                        right: 14,
                        bottom: isPrintMode() ? 24 : 0,
                        left: 14
                    }
                },
                annotations: { yaxis: [] }
            };

            const flowOptions = {
                chart: { type: 'line', height: isPrintMode() ? 390 : 420, width: '100%', parentHeightOffset: 0, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, selection: { enabled: false }, dropShadow: { enabled: false }, animations: { enabled: false }, zoom: { enabled: false }, selection: { enabled: false }, dropShadow: { enabled: false }, sparkline: { enabled: false } },
                series: [
                    { name: 'Entrada', data: smoothFlowSeries(sim.map(s => Number(s.entrada)), 8) },
                    { name: 'Salida', data: smoothFlowSeries(sim.map(s => Number(s.salidaTotal)), 8) }
                ],
                colors: ['#10b981', '#e11d48'],
                stroke: { curve: 'smooth', width: [3, 3], lineCap: 'round' },
                markers: {
                    size: 0,
                    strokeWidth: 2,
                    strokeColors: '#ffffff',
                    hover: {
                        size: 7,
                        sizeOffset: 3
                    }
                },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: isPrintMode() ? 6 : 10, tickPlacement: 'between', labels: { rotate: isPrintMode() ? 0 : -90, rotateAlways: !isPrintMode(), hideOverlappingLabels: true, trim: false, formatter: value => isPrintMode() && value ? String(value).replace(' ', '\n') : value, style: { fontSize: isPrintMode() ? '6px' : '10px', colors: '#64748b' }, offsetY: isPrintMode() ? 4 : 0 }, tooltip: { enabled: false } },
                yaxis: { labels: { formatter: v => v.toFixed(1), style: { colors: '#64748b' } } },
                tooltip: {
                    enabled: true,
                    shared: true,
                    intersect: false,
                    followCursor: true,
                    hideEmptySeries: true,
                    marker: {
                        show: true
                    },
                    x: {
                        show: true
                    },
                    y: {
                        formatter: value => {
                            if (value === undefined || value === null) {
                                return '';
                            }

                            return `${value.toFixed(2)} l/s`;
                        }
                    }
                },
                legend: {
                    show: true,
                    position: 'bottom',
                    horizontalAlign: 'center',
                    floating: false,
                    fontSize: isPrintMode() ? '8px' : '12px',
                    fontWeight: 600,
                    offsetX: 0,
                    offsetY: isPrintMode() ? 10 : 0,
                    itemMargin: {
                        horizontal: 10,
                        vertical: 0
                    }
                },
                grid: {
                    borderColor: '#e2e8f0',
                    padding: {
                        top: 5,
                        right: 14,
                        bottom: isPrintMode() ? 24 : 0,
                        left: 14
                    }
                },
            };

            if (chartEntNivel) chartEntNivel.destroy();
            chartEntNivel = new ApexCharts(document.querySelector('#chart-entronque-nivel'), levelOptions);
            chartEntNivel.render();

            if (chartEntCaudales) chartEntCaudales.destroy();
            chartEntCaudales = new ApexCharts(document.querySelector('#chart-entronque-caudales'), flowOptions);
            chartEntCaudales.render();
        }

        // Modal Functions for Excel Data
        function openExcelModal(target) {
            activeModalTarget = target;
            const targetObj = target === 'alcala' ? state.alcala : state.entronque;

            tempHourlyProfile = {
                entrada: [...targetObj.entradaProfile],
                salida1: [...targetObj.salida1Profile],
                burguillos: target === 'alcala' ? [...targetObj.burguillosProfile] : Array(24).fill(0)
            };

            document.getElementById('excel-modal-title').innerText = `Cargar / Pegar Datos de Excel - ${target === 'alcala' ? 'Alcalá del Río' : 'Entronque'}`;
            document.getElementById('modal-th-burguillos').style.display = target === 'alcala' ? '' : 'none';
            document.getElementById('excel-paste-area').value = '';

            renderModalTable();
            document.getElementById('excel-modal').classList.remove('hidden');
        }

        function closeExcelModal() {
            document.getElementById('excel-modal').classList.add('hidden');
        }

        function renderModalTable() {
            const tbody = document.getElementById('modal-table-body');
            let html = '';

            for (let h = 0; h < 24; h++) {
                const timeStr = `${String(h).padStart(2, '0')}:00`;
                html += `
                    <tr>
                        <td class="py-1.5 px-4 font-semibold text-slate-700">${timeStr}</td>
                        <td class="py-1.5 px-4">
                            <input type="number" value="${tempHourlyProfile.entrada[h]}" onchange="tempHourlyProfile.entrada[${h}] = parseFloat(this.value)||0" class="w-24 px-2 py-1 border border-slate-300 rounded text-xs font-semibold focus:ring-1 focus:ring-emerald-500">
                        </td>
                        <td class="py-1.5 px-4">
                            <input type="number" value="${tempHourlyProfile.salida1[h]}" onchange="tempHourlyProfile.salida1[${h}] = parseFloat(this.value)||0" class="w-24 px-2 py-1 border border-slate-300 rounded text-xs focus:ring-1 focus:ring-emerald-500">
                        </td>
                        ${activeModalTarget === 'alcala' ? `
                            <td class="py-1.5 px-4">
                                <input type="number" value="${tempHourlyProfile.burguillos[h]}" onchange="tempHourlyProfile.burguillos[${h}] = parseFloat(this.value)||0" class="w-24 px-2 py-1 border border-slate-300 rounded text-xs focus:ring-1 focus:ring-emerald-500">
                            </td>
                        ` : ''}
                    </tr>
                `;
            }
            tbody.innerHTML = html;
        }

        // Process pasted column text from Excel
        function processPastedExcelData() {
            const rawText = document.getElementById('excel-paste-area').value.trim();
            if (!rawText) return;

            // Extract numeric values from pasted content
            const numbers = rawText
                .split(/[\n,;\t]+/)
                .map(v => parseFloat(v.replace(/[^0-9.-]/g, '')))
                .filter(v => !isNaN(v));

            if (numbers.length === 0) return;

            for (let i = 0; i < 24; i++) {
                if (i < numbers.length) {
                    tempHourlyProfile.entrada[i] = numbers[i];
                } else {
                    tempHourlyProfile.entrada[i] = numbers[numbers.length - 1]; // pad last if fewer than 24
                }
            }

            renderModalTable();
        }

        // Handle Excel File Upload (.xlsx, .xls, .csv)
        function handleExcelFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    const parsedNumbers = [];
                    json.forEach(row => {
                        row.forEach(cell => {
                            const val = parseFloat(cell);
                            if (!isNaN(val) && val > 0 && val < 10000) {
                                parsedNumbers.push(val);
                            }
                        });
                    });

                    if (parsedNumbers.length >= 24) {
                        for (let i = 0; i < 24; i++) {
                            tempHourlyProfile.entrada[i] = parsedNumbers[i];
                        }
                    } else if (parsedNumbers.length > 0) {
                        for (let i = 0; i < 24; i++) {
                            tempHourlyProfile.entrada[i] = parsedNumbers[i % parsedNumbers.length];
                        }
                    }

                    renderModalTable();
                } catch (err) {
                    console.error('Error al leer Excel:', err);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function resetToDefaultProfile() {
            if (activeModalTarget === 'alcala') {
                tempHourlyProfile.entrada = [...PATTERN_ENTRADA_ALCALA_DEFAULT];
                tempHourlyProfile.salida1 = [...PATTERN_SALIDA_1_ALCALA];
                tempHourlyProfile.burguillos = [...PATTERN_BURGUILLOS];
            } else {
                tempHourlyProfile.entrada = [...PATTERN_ENTRADA_ENTRONQUE_DEFAULT];
                tempHourlyProfile.salida1 = [...PATTERN_SALIDA_1_ENTRONQUE];
            }
            renderModalTable();
        }

        function saveModalChanges() {
            const targetObj = activeModalTarget === 'alcala' ? state.alcala : state.entronque;
            targetObj.entradaProfile = [...tempHourlyProfile.entrada];
            targetObj.salida1Profile = [...tempHourlyProfile.salida1];
            if (activeModalTarget === 'alcala') {
                targetObj.burguillosProfile = [...tempHourlyProfile.burguillos];
            }
            targetObj.isCustomData = true;

            const newAvg = Math.round(targetObj.entradaProfile.reduce((a,b)=>a+b, 0) / 24);
            targetObj.caudalEntradaMedio = newAvg;

            if (activeModalTarget === 'alcala') {
                document.getElementById('alc-caudal-entrada').value = newAvg;
                updateAlcalaSimulation();
            } else {
                document.getElementById('ent-caudal-entrada').value = newAvg;
                updateEntronqueSimulation();
            }

            closeExcelModal();
        }

        function switchTab(tab) {
            state.currentTab = tab;
            const isEntronque = tab === 'entronque';
            document.getElementById('tab-alcala').classList.toggle('hidden', isEntronque);
            document.getElementById('tab-entronque').classList.toggle('hidden', !isEntronque);
            document.getElementById('tab-btn-alcala').className = !isEntronque ? 'py-3 px-1 tab-active flex items-center space-x-2' : 'py-3 px-1 text-slate-400 hover:text-white transition-all flex items-center space-x-2';
            document.getElementById('tab-btn-entronque').className = isEntronque ? 'py-3 px-1 tab-active flex items-center space-x-2' : 'py-3 px-1 text-slate-400 hover:text-white transition-all flex items-center space-x-2';
            if (isEntronque) updateEntronqueUI(); else updateAlcalaUI();
        }

        function setBurguillosOption(opt) {
            state.alcala.burguillosOption = opt;
            const btnOn = document.getElementById('btn-burguillos-on');
            const btnOff = document.getElementById('btn-burguillos-off');

            if (opt === 'con') {
                btnOn.className = 'px-3 py-1.5 text-xs font-medium rounded-lg bg-white text-slate-800 shadow-sm transition-all';
                btnOff.className = 'px-3 py-1.5 text-xs font-medium rounded-lg text-slate-600 hover:text-slate-900 transition-all';
            } else {
                btnOff.className = 'px-3 py-1.5 text-xs font-medium rounded-lg bg-white text-slate-800 shadow-sm transition-all';
                btnOn.className = 'px-3 py-1.5 text-xs font-medium rounded-lg text-slate-600 hover:text-slate-900 transition-all';
            }
            updateAlcalaSimulation();
        }

        function updateAlcalaSimulation() {
            state.alcala.nivelInicio = parseFloat(document.getElementById('alc-nivel-inicio').value) || 0;
            state.alcala.nivelMinimo = parseFloat(document.getElementById('alc-nivel-minimo').value) || 0;
            updateAlcalaUI();
        }

        function updateEntronqueSimulation() {
            state.entronque.nivelInicio = parseFloat(document.getElementById('ent-nivel-inicio').value) || 0;
            state.entronque.nivelMinimo = parseFloat(document.getElementById('ent-nivel-minimo').value) || 0;
            updateEntronqueUI();
        }

        function exportSummaryPDF() {
            window.print();
        }

        function normalizeSheetName(name) {
            return String(name || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
        }

        function applyWorkbookSheet(workbook, target) {
            const wanted = target === 'alcala' ? 'alcala' : 'entronque';
            const sheetName = workbook.SheetNames.find(name => normalizeSheetName(name).includes(wanted));
            if (!sheetName) return false;
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
            const firstDate = rows.slice(1).map(row => normalizeExcelDate(row[0])).find(Boolean);
            const records = rows.slice(1).map(row => ({
                date: normalizeExcelDate(row[0]),
                entrada: Number(row[1]),
                salida1: Number(row[2]),
                burguillos: target === 'alcala' ? Number(row[3] || 0) : 0
            })).filter(record => record.date && Number.isFinite(record.entrada) && Number.isFinite(record.salida1));
            if (!firstDate || !records.length) return false;
            records.sort((a, b) => a.date - b.date);
            const targetObj = state[target];
            targetObj.startDate = new Date(firstDate.getTime());
            targetObj.startDate.setMilliseconds(0);
            syncManeuverDateInput(target);
            syncInletAverageField(target);
            targetObj.entradaProfile = records.map(record => record.entrada);
            targetObj.salida1Profile = records.map(record => record.salida1);
            if (target === 'alcala') targetObj.burguillosProfile = records.map(record => record.burguillos);
            targetObj.timestamps = Array.from({ length: Math.min(168, records.length) }, (_, i) => new Date(targetObj.startDate.getTime() + i * 3600000));
            targetObj.maxHours = Math.min(168, records.length);
            targetObj.isCustomData = true;
            const average = targetObj.entradaProfile.reduce((sum, value) => sum + value, 0) / targetObj.entradaProfile.length;
            targetObj.caudalEntradaMedio = average;
            const input = document.getElementById(target === 'alcala' ? 'alc-caudal-entrada' : 'ent-caudal-entrada');
            if (input) input.value = getInletProfileAverage(target).toFixed(2);
            syncInletAverageField(target);
            return true;
        }

        async function loadDefaultExcel() {
            try {
                const response = await fetch('./Consumo_Patron.xlsx', { cache: 'no-store' });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);
                const workbook = XLSX.read(await response.arrayBuffer(), {
                    type: 'array',
                    cellDates: true
                });
                const loadedAlcala = applyWorkbookSheet(workbook, 'alcala');
                const loadedEntronque = applyWorkbookSheet(workbook, 'entronque');
                syncManeuverDateInput('alcala');
                syncManeuverDateInput('entronque');
                syncInletAverageField('alcala');
                syncInletAverageField('entronque');
                if (!loadedAlcala && !loadedEntronque) {
                    throw new Error('No se encontraron las hojas esperadas.');
                }
                updateAlcalaUI();
                updateEntronqueUI();
            } catch (error) {
                console.warn('No se pudo cargar Consumo_Patron.xlsx. Se mantienen los perfiles de respaldo.', error);
            }
        }

        window.onload = async function() {
            await loadDefaultExcel();
            updateAlcalaUI();
            updateEntronqueUI();
            switchTab('alcala');
        };