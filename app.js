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
        let chartResNiveles = null;
        let chartResDistribucion = null;

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
                nivelInicio: 15.50,
                nivelMinimo: 12.00,
                caudalEntradaMedio: 44.36,
                isCustomData: false,
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
                nivelInicio: 18.20,
                nivelMinimo: 14.00,
                caudalEntradaMedio: 48.50,
                isCustomData: false,
                entradaProfile: [...PATTERN_ENTRADA_ENTRONQUE_DEFAULT],
                salida1Profile: [...PATTERN_SALIDA_1_ENTRONQUE],
                factor: 1798,
                maxHours: 168,
                visibleHours: 168,
                breachHour: -1,
                simulation: []
            }
        };

        function getTimeLabels(days) {
            const labels = [];
            for (let d = 1; d <= days; d++) {
                for (let h = 0; h < 24; h++) {
                    const timeStr = `${String(h).padStart(2, '0')}:00`;
                    labels.push(days === 1 ? timeStr : `Día ${d} ${timeStr}`);
                }
            }
            return labels;
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
                    const avg = (state.alcala.entradaProfile.reduce((a,b)=>a+b, 0) / 24).toFixed(2);
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
                    const avg = (state.entronque.entradaProfile.reduce((a,b)=>a+b, 0) / 24).toFixed(2);
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
            const maxHours = state.alcala.maxHours;
            const labels = getTimeLabels(Math.ceil(maxHours / 24));
            let breachHour = -1;

            for (let i = 0; i < maxHours; i++) {
                const hourOfDay = i % 24;
                const dayNum = Math.floor(i / 24) + 1;
                const entrada = useExcelProfile ? (state.alcala.entradaProfile[hourOfDay] ?? 44.36) : fixedInletVal;
                const salida1 = state.alcala.salida1Profile[hourOfDay] ?? 29.9;
                const burguillos = isBurguillosActive ? (state.alcala.burguillosProfile[hourOfDay] ?? 0) : 0;
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
            const maxHours = state.entronque.maxHours;
            const labels = getTimeLabels(Math.ceil(maxHours / 24));
            let breachHour = -1;

            for (let i = 0; i < maxHours; i++) {
                const hourOfDay = i % 24;
                const dayNum = Math.floor(i / 24) + 1;
                const entrada = useExcelProfile ? (state.entronque.entradaProfile[hourOfDay] ?? 48.5) : fixedInletVal;
                const salida1 = state.entronque.salida1Profile[hourOfDay] ?? 30.0;
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

        // Update UI for Alcalá
        function updateAlcalaUI() {
            calculateAlcala();
            const sim = state.alcala.simulation;
            const minNivelReq = parseFloat(state.alcala.nivelMinimo);

            const nivelFinal = sim[sim.length - 1].nivel;
            let minIndex = 0;
            let nivelMinAlcanzado = sim[0].nivel;
            for (let i = 1; i < sim.length; i++) {
                if (sim[i].nivel < nivelMinAlcanzado) {
                    nivelMinAlcanzado = sim[i].nivel;
                    minIndex = i;
                }
            }

            const entradaAvg = Math.round(sim.reduce((acc, s) => acc + s.entrada, 0) / sim.length);
            const salidaTotalAvg = Math.round(sim.reduce((acc, s) => acc + s.salidaTotal, 0) / sim.length);
            const burguillosAvg = Math.round(sim.reduce((acc, s) => acc + s.burguillos, 0) / sim.length);
            const diffNivel = nivelFinal - state.alcala.nivelInicio;

            document.getElementById('alc-label-nivel-final').innerText = 'Nivel Final';
            document.getElementById('alc-kpi-nivel-final').innerText = `${nivelFinal.toFixed(2)} m`;
            
            const diffEl = document.getElementById('alc-kpi-diff-nivel');
            diffEl.innerText = `${diffNivel >= 0 ? '+' : ''}${diffNivel.toFixed(2)} m vs inicio`;
            diffEl.className = `text-xs font-semibold ${diffNivel >= 0 ? 'text-emerald-600' : 'text-rose-600'}`;

            const minEl = document.getElementById('alc-kpi-nivel-min');
            minEl.innerText = `${nivelMinAlcanzado.toFixed(2)} m`;
            minEl.className = `text-2xl font-bold mt-1 ${nivelMinAlcanzado < minNivelReq ? 'text-rose-600' : 'text-slate-800'}`;

            const statusMinEl = document.getElementById('alc-kpi-status-min');
            if (nivelMinAlcanzado < minNivelReq) {
                const minTimeLabel = sim[minIndex].hora;
                statusMinEl.innerText = `¡ALERTA MÍNIMO! (${minTimeLabel})`;
                statusMinEl.className = 'text-xs font-bold text-rose-600 animate-pulse';
            } else {
                statusMinEl.innerText = 'Nivel Seguro en Horizonte';
                statusMinEl.className = 'text-xs font-semibold text-emerald-600';
            }

            document.getElementById('alc-kpi-horizon-text').innerText = `Promedio en ${state.alcala.visibleHours} h`;
            document.getElementById('alc-kpi-entrada-avg').innerText = `${entradaAvg} l/s`;
            document.getElementById('alc-kpi-salida-total').innerText = `${salidaTotalAvg} l/s`;
            document.getElementById('alc-kpi-burguillos-status').innerText = `Burguillos medio: ${state.alcala.burguillosOption === 'con' ? burguillosAvg + ' l/s' : '0 l/s'}`;

            updateAlcalaCharts();
        }

        function smoothFlowSeries(values, radius = 3) {
            if (!Array.isArray(values) || values.length === 0) return [];
            return values.map((value, index) => {
                let weightedSum = 0;
                let weightTotal = 0;
                for (let offset = -radius; offset <= radius; offset++) {
                    const sourceIndex = Math.min(values.length - 1, Math.max(0, index + offset));
                    const weight = radius + 1 - Math.abs(offset);
                    weightedSum += values[sourceIndex] * weight;
                    weightTotal += weight;
                }
                return Number((weightedSum / weightTotal).toFixed(2));
            });
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

            document.getElementById('alc-chart-max').innerText = `Máx: ${maxValue.toFixed(2)} m (${labels[maxIndex]})`;
            document.getElementById('alc-chart-min').innerText = `Mín: ${minValue.toFixed(2)} m (${labels[minIndex]})`;

            const levelOptions = {
                chart: { type: 'line', height: 380, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, animations: { enabled: false }, zoom: { enabled: false } },
                series: [{ name: 'Nivel del depósito', data: niveles }],
                colors: ['#0284c7'],
                stroke: { curve: 'smooth', width: 3 },
                markers: { size: 0, hover: { size: 5 } },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: 10, labels: { rotate: -45, hideOverlappingLabels: true, style: { fontSize: '10px', colors: '#64748b' } }, tooltip: { enabled: false } },
                yaxis: { min: Math.max(0, Math.floor(Math.min(minValue, minLimit) - 0.5)), max: Math.ceil(maxValue + 0.5), labels: { formatter: v => `${v.toFixed(2)} m`, style: { colors: '#64748b' } } },
                tooltip: { shared: false, intersect: false, followCursor: true, x: { show: true }, y: { formatter: v => `${v.toFixed(2)} m` } },
                legend: { show: false },
                grid: { borderColor: '#e2e8f0' },
                annotations: { yaxis: [{ y: minLimit, borderColor: '#ef4444', strokeDashArray: 5 }] }
            };

            const flowOptions = {
                chart: { type: 'line', height: 380, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, animations: { enabled: false }, zoom: { enabled: false } },
                series: [
                    { name: 'Entrada', data: smoothFlowSeries(sim.map(s => Number(s.entrada)), 3) },
                    { name: 'Salida (Salida 1 + Burguillos)', data: smoothFlowSeries(sim.map(s => Number(s.salidaTotal)), 3) }
                ],
                colors: ['#10b981', '#e11d48'],
                stroke: { curve: 'smooth', width: [3, 3], lineCap: 'round' },
                markers: { size: 0, hover: { size: 5 } },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: 10, labels: { rotate: -45, hideOverlappingLabels: true, style: { fontSize: '10px', colors: '#64748b' } }, tooltip: { enabled: false } },
                yaxis: { labels: { formatter: v => `${v.toFixed(1)} l/s`, style: { colors: '#64748b' } } },
                tooltip: { shared: false, intersect: false, followCursor: true, y: { formatter: v => `${v.toFixed(2)} l/s` } },
                legend: { position: 'top', horizontalAlign: 'center', fontSize: '12px', fontWeight: 600 },
                grid: { borderColor: '#e2e8f0' }
            };

            if (chartAlcNivel) chartAlcNivel.destroy();
            chartAlcNivel = new ApexCharts(document.querySelector('#chart-alcala-nivel'), levelOptions);
            chartAlcNivel.render();

            if (chartAlcCaudales) chartAlcCaudales.destroy();
            chartAlcCaudales = new ApexCharts(document.querySelector('#chart-alcala-caudales'), flowOptions);
            chartAlcCaudales.render();
        }

        // Update UI for Entronque
        function updateEntronqueUI() {
            calculateEntronque();
            const sim = state.entronque.simulation;
            const minNivelReq = parseFloat(state.entronque.nivelMinimo);

            const nivelFinal = sim[sim.length - 1].nivel;
            let minIndex = 0;
            let nivelMinAlcanzado = sim[0].nivel;
            for (let i = 1; i < sim.length; i++) {
                if (sim[i].nivel < nivelMinAlcanzado) {
                    nivelMinAlcanzado = sim[i].nivel;
                    minIndex = i;
                }
            }

            const entradaAvg = Math.round(sim.reduce((acc, s) => acc + s.entrada, 0) / sim.length);
            const salidaAvg = Math.round(sim.reduce((acc, s) => acc + s.salida1, 0) / sim.length);
            const diffNivel = nivelFinal - state.entronque.nivelInicio;

            document.getElementById('ent-kpi-nivel-final').innerText = `${nivelFinal.toFixed(2)} m`;
            
            const diffEl = document.getElementById('ent-kpi-diff-nivel');
            diffEl.innerText = `${diffNivel >= 0 ? '+' : ''}${diffNivel.toFixed(2)} m vs inicio`;
            diffEl.className = `text-xs font-semibold ${diffNivel >= 0 ? 'text-emerald-600' : 'text-rose-600'}`;

            const minEl = document.getElementById('ent-kpi-nivel-min');
            minEl.innerText = `${nivelMinAlcanzado.toFixed(2)} m`;
            minEl.className = `text-2xl font-bold mt-1 ${nivelMinAlcanzado < minNivelReq ? 'text-rose-600' : 'text-slate-800'}`;

            const statusMinEl = document.getElementById('ent-kpi-status-min');
            if (nivelMinAlcanzado < minNivelReq) {
                const minTimeLabel = sim[minIndex].hora;
                statusMinEl.innerText = `¡ALERTA MÍNIMO! (${minTimeLabel})`;
                statusMinEl.className = 'text-xs font-bold text-rose-600 animate-pulse';
            } else {
                statusMinEl.innerText = 'Nivel Seguro en Horizonte';
                statusMinEl.className = 'text-xs font-semibold text-emerald-600';
            }

            document.getElementById('ent-kpi-entrada-avg').innerText = `${entradaAvg} l/s`;
            document.getElementById('ent-kpi-salida-avg').innerText = `${salidaAvg} l/s`;

            updateEntronqueCharts();
        }

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

            document.getElementById('ent-chart-max').innerText = `Máx: ${maxValue.toFixed(2)} m (${labels[maxIndex]})`;
            document.getElementById('ent-chart-min').innerText = `Mín: ${minValue.toFixed(2)} m (${labels[minIndex]})`;

            const levelOptions = {
                chart: { type: 'line', height: 380, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, animations: { enabled: false }, zoom: { enabled: false } },
                series: [{ name: 'Nivel del depósito', data: niveles }],
                colors: ['#6366f1'],
                stroke: { curve: 'smooth', width: 3 },
                markers: { size: 0, hover: { size: 5 } },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: 10, labels: { rotate: -45, hideOverlappingLabels: true, style: { fontSize: '10px', colors: '#64748b' } }, tooltip: { enabled: false } },
                yaxis: { min: Math.max(0, Math.floor(Math.min(minValue, minLimit) - 0.5)), max: Math.ceil(maxValue + 0.5), labels: { formatter: v => `${v.toFixed(2)} m`, style: { colors: '#64748b' } } },
                tooltip: { shared: false, intersect: false, followCursor: true, x: { show: true }, y: { formatter: v => `${v.toFixed(2)} m` } },
                legend: { show: false },
                grid: { borderColor: '#e2e8f0' },
                annotations: { yaxis: [{ y: minLimit, borderColor: '#ef4444', strokeDashArray: 5 }] }
            };

            const flowOptions = {
                chart: { type: 'line', height: 380, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, animations: { enabled: false }, zoom: { enabled: false } },
                series: [
                    { name: 'Entrada', data: smoothFlowSeries(sim.map(s => Number(s.entrada)), 3) },
                    { name: 'Salida', data: smoothFlowSeries(sim.map(s => Number(s.salidaTotal)), 3) }
                ],
                colors: ['#10b981', '#e11d48'],
                stroke: { curve: 'smooth', width: [3, 3], lineCap: 'round' },
                markers: { size: 0, hover: { size: 5 } },
                dataLabels: { enabled: false },
                xaxis: { categories: labels, tickAmount: 10, labels: { rotate: -45, hideOverlappingLabels: true, style: { fontSize: '10px', colors: '#64748b' } }, tooltip: { enabled: false } },
                yaxis: { labels: { formatter: v => `${v.toFixed(1)} l/s`, style: { colors: '#64748b' } } },
                tooltip: { shared: false, intersect: false, followCursor: true, y: { formatter: v => `${v.toFixed(2)} l/s` } },
                legend: { position: 'top', horizontalAlign: 'center', fontSize: '12px', fontWeight: 600 },
                grid: { borderColor: '#e2e8f0' }
            };

            if (chartEntNivel) chartEntNivel.destroy();
            chartEntNivel = new ApexCharts(document.querySelector('#chart-entronque-nivel'), levelOptions);
            chartEntNivel.render();

            if (chartEntCaudales) chartEntCaudales.destroy();
            chartEntCaudales = new ApexCharts(document.querySelector('#chart-entronque-caudales'), flowOptions);
            chartEntCaudales.render();
        }

        function updateResumenUI() {
            calculateAlcala();
            calculateEntronque();

            const simAlc = state.alcala.simulation;
            const simEnt = state.entronque.simulation;

            const alcMin = Math.min(...simAlc.map(s => s.nivel));
            const alcMinLimit = parseFloat(state.alcala.nivelMinimo);
            const alcMargen = alcMin - alcMinLimit;
            const alcFin = simAlc[simAlc.length - 1].nivel;
            const alcVolumenNeto = simAlc.reduce((acc, s) => acc + (s.netFlow * 3.6), 0);

            const entMin = Math.min(...simEnt.map(s => s.nivel));
            const entMinLimit = parseFloat(state.entronque.nivelMinimo);
            const entMargen = entMin - entMinLimit;
            const entFin = simEnt[simEnt.length - 1].nivel;
            const entVolumenNeto = simEnt.reduce((acc, s) => acc + (s.netFlow * 3.6), 0);

            const alcOk = alcMin >= alcMinLimit;
            const entOk = entMin >= entMinLimit;

            document.getElementById('res-badge-alc').className = `text-xs px-2.5 py-0.5 rounded-full font-semibold ${alcOk ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}`;
            document.getElementById('res-badge-alc').innerText = alcOk ? 'ÓPTIMO' : 'ALERTA MÍNIMO';

            document.getElementById('res-badge-ent').className = `text-xs px-2.5 py-0.5 rounded-full font-semibold ${entOk ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}`;
            document.getElementById('res-badge-ent').innerText = entOk ? 'ÓPTIMO' : 'ALERTA MÍNIMO';

            const burguillosAvg = Math.round(simAlc.reduce((acc, s) => acc + s.burguillos, 0) / simAlc.length);
            document.getElementById('res-alc-burguillos').innerText = state.alcala.burguillosOption === 'con' ? `Con Burguillos (~${burguillosAvg} l/s)` : 'Sin Burguillos (0 l/s)';
            document.getElementById('res-alc-niveles').innerText = `${parseFloat(state.alcala.nivelInicio).toFixed(2)}m / ${alcFin.toFixed(2)}m`;
            document.getElementById('res-alc-min').innerText = `${alcMin.toFixed(2)} m`;
            document.getElementById('res-alc-margen').innerText = `${alcMargen >= 0 ? '+' : ''}${alcMargen.toFixed(2)} m`;
            document.getElementById('res-alc-margen').className = `font-bold ${alcMargen >= 0 ? 'text-emerald-600' : 'text-rose-600'}`;
            document.getElementById('res-alc-volumen').innerText = `${alcVolumenNeto >= 0 ? '+' : ''}${Math.round(alcVolumenNeto).toLocaleString()} m³`;

            document.getElementById('res-ent-niveles').innerText = `${parseFloat(state.entronque.nivelInicio).toFixed(2)}m / ${entFin.toFixed(2)}m`;
            document.getElementById('res-ent-min').innerText = `${entMin.toFixed(2)} m`;
            document.getElementById('res-ent-margen').innerText = `${entMargen >= 0 ? '+' : ''}${entMargen.toFixed(2)} m`;
            document.getElementById('res-ent-margen').className = `font-bold ${entMargen >= 0 ? 'text-emerald-600' : 'text-rose-600'}`;
            document.getElementById('res-ent-volumen').innerText = `${entVolumenNeto >= 0 ? '+' : ''}${Math.round(entVolumenNeto).toLocaleString()} m³`;

            const globalStatusEl = document.getElementById('resumen-global-status');
            if (alcOk && entOk) {
                globalStatusEl.innerText = 'Sistema Estable (Sin Riesgos)';
                globalStatusEl.className = 'text-sm font-bold text-emerald-400';
            } else {
                globalStatusEl.innerText = 'Atención Requerida (Nivel Crítico)';
                globalStatusEl.className = 'text-sm font-bold text-rose-400 animate-pulse';
            }

            const recEl = document.getElementById('resumen-recomendaciones');
            let htmlRec = '';

            if (!alcOk) {
                htmlRec += `
                    <div class="p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-800 flex items-start space-x-2">
                        <i class="fa-solid fa-triangle-exclamation text-rose-500 mt-0.5"></i>
                        <div><strong>Alcalá del Río:</strong> El nivel desciende por debajo del mínimo de seguridad (${alcMinLimit}m). Se sugiere incrementar caudal de entrada o desactivar Burguillos.</div>
                    </div>`;
            } else {
                htmlRec += `
                    <div class="p-2.5 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-800 flex items-start space-x-2">
                        <i class="fa-solid fa-circle-check text-emerald-500 mt-0.5"></i>
                        <div><strong>Alcalá del Río:</strong> Operando con margen de seguridad óptimo (${alcMargen.toFixed(2)}m sobre el mínimo).</div>
                    </div>`;
            }

            if (!entOk) {
                htmlRec += `
                    <div class="p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-800 flex items-start space-x-2">
                        <i class="fa-solid fa-triangle-exclamation text-rose-500 mt-0.5"></i>
                        <div><strong>Entronque:</strong> Nivel por debajo de cota de seguridad (${entMinLimit}m). Elevar aportación de consigna en la entrada.</div>
                    </div>`;
            } else {
                htmlRec += `
                    <div class="p-2.5 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-800 flex items-start space-x-2">
                        <i class="fa-solid fa-circle-check text-emerald-500 mt-0.5"></i>
                        <div><strong>Entronque:</strong> Cota de reserva estable (${entMargen.toFixed(2)}m sobre el mínimo).</div>
                    </div>`;
            }

            recEl.innerHTML = htmlRec;

            // Las gráficas principales usan ApexCharts. El resumen conserva sus indicadores y recomendaciones.
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

            document.getElementById('tab-alcala').classList.add('hidden');
            document.getElementById('tab-entronque').classList.add('hidden');
            document.getElementById('tab-resumen').classList.add('hidden');

            document.getElementById('tab-btn-alcala').className = 'py-3 px-1 text-slate-400 hover:text-white transition-all flex items-center space-x-2';
            document.getElementById('tab-btn-entronque').className = 'py-3 px-1 text-slate-400 hover:text-white transition-all flex items-center space-x-2';
            document.getElementById('tab-btn-resumen').className = 'py-3 px-1 text-slate-400 hover:text-white transition-all flex items-center space-x-2';

            if (tab === 'alcala') {
                document.getElementById('tab-alcala').classList.remove('hidden');
                document.getElementById('tab-btn-alcala').className = 'py-3 px-1 tab-active flex items-center space-x-2';
                updateAlcalaUI();
            } else if (tab === 'entronque') {
                document.getElementById('tab-entronque').classList.remove('hidden');
                document.getElementById('tab-btn-entronque').className = 'py-3 px-1 tab-active flex items-center space-x-2';
                updateEntronqueUI();
            } else if (tab === 'resumen') {
                document.getElementById('tab-resumen').classList.remove('hidden');
                document.getElementById('tab-btn-resumen').className = 'py-3 px-1 tab-active flex items-center space-x-2';
                updateResumenUI();
            }
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

        window.onload = function() {
            updateAlcalaUI();
            updateEntronqueUI();
            switchTab('alcala');
        };
