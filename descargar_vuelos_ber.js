// Pegar en la consola de DevTools estando en https://ber.berlin-airport.de/en/flying/departures-arrivals.html
// Pide el listado de salidas y llegadas de cada fecha (las mismas consultas que hace la página)
// y descarga un archivo vuelos_ber_api.json con número de vuelo, horario, modelo y matrícula.
(async () => {
  const dias = ['2026-10-22', '2026-10-29'];
  const out = [];
  for (const d of dias) {
    const sig = new Date(d + 'T12:00:00Z'); sig.setUTCDate(sig.getUTCDate() + 1);
    const hasta = sig.toISOString().slice(0, 10);
    for (const dir of ['D', 'A']) {
      let page = 1, total = 1;
      do {
        const url = `/api.flights.json?arrivalDeparture=${dir}&dateFrom=${d}T00:00:00&dateUntil=${hasta}` +
                    `&search=&lang=en&page=${page}&terminal=&itemsPerPage=50`;
        const r = await fetch(url, { headers: { Accept: 'application/json' } });
        if (!r.ok) { console.warn(d, dir, 'página', page, 'error', r.status); break; }
        const data = (await r.json()).data || {};
        for (const it of data.items || []) {
          out.push({ flight_number: it.flight_number, is_departure: it.is_departure,
                     scheduled_time: it.scheduled_time, aircraft_type: it.aircraft_type,
                     aircraft_reg: it.aircraft_reg });
        }
        total = data.total_pages || 1;
        console.log(d, dir === 'D' ? 'salidas' : 'llegadas', `página ${page}/${total}`);
        page++;
        await new Promise(res => setTimeout(res, 700)); // pausa para no cargar el sitio
      } while (page <= total);
    }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify(out)], { type: 'application/json' }));
  a.download = 'vuelos_ber_api.json';
  a.click();
  console.log('Listo:', out.length, 'vuelos descargados');
})();
