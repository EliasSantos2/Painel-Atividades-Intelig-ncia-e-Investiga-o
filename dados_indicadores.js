// Gerado automaticamente por extrair_indicadores.ps1
// Data: 2026-07-06T14:25:47
// NAO editar manualmente.

const DADOS_IND = {
  meses: ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'],
  trechos: ['Trecho 0','Trecho 1','Trecho 2','Trecho 3'],
  anos: [2024, 2025, 2026],
  geradoEm: '2026-07-06T14:25:47',
  data: {
    '2024': {
      't0': { circ:[1076,1359,1592,1454,1546,1518,1536,1484,1519,1575,1480,1205], vand:[33,79,62,39,60,60,75,74,53,46,30,70], thp:[31.02,77.85,33.62,31.3,52.75,37.63,32.95,32.22,28.22,37.1,12.85,46.52], vagab:[8,11,9,9,19,15,17,8,7,10,2,19], txab:[0.04,0.04,0.03,0.03,0.06,0.04,0.05,0.02,0.02,0.03,0.01,0.08] },
      't1': { circ:[1975,2189,2477,2004,2455,2351,2397,2265,2305,2487,2444,2014], vand:[20,22,17,16,15,16,10,24,32,12,33,28], thp:[39.62,20.72,23.18,21.35,19.77,27,3.42,5.15,10.23,8.45,19.67,9.57], vagab:[0,0,0,0,0,0,0,2,1,0,0,0], txab:[null,null,null,null,null,null,null,null,null,null,null,null] },
      't2': { circ:[null,null,null,null,null,null,null,null,null,null,null,null], vand:[null,null,null,null,null,null,null,null,null,null,null,null], thp:[null,null,null,null,null,null,null,null,null,null,null,null], vagab:[null,null,null,null,null,null,null,null,null,null,null,null], txab:[null,null,null,null,null,null,null,null,null,null,null,null] },
      't3': { circ:[null,null,null,null,null,null,null,null,null,null,null,null], vand:[null,null,null,null,null,null,null,null,null,null,null,null], thp:[null,null,null,null,null,null,null,null,null,null,null,null], vagab:[null,null,null,null,null,null,null,null,null,null,null,null], txab:[null,null,null,null,null,null,null,null,null,null,null,null] }
    },
    '2025': {
      't0': { circ:[771,1098,1526,1469,1609,1594,1602,1694,1620,1577,1567,1550], vand:[48,47,59,91,80,40,28,21,27,44,27,18], thp:[45.67,45.03,41.38,75.02,51.13,53.28,24.13,22.58,31.37,67.58,49.32,38.57], vagab:[14,6,15,39,20,10,18,12,10,24,6,8], txab:[0.11,0.02,0.05,0.12,0.05,0.03,0.06,0.03,0.03,0.07,0.02,0.02] },
      't1': { circ:[1338,1751,2156,2161,2233,2293,2479,2472,2628,2477,2456,2440], vand:[24,32,26,15,12,15,14,17,21,16,21,20], thp:[39.57,55.23,37.38,22.83,23.02,28.5,24.32,27.37,30.43,29.55,36.55,33.33], vagab:[0,0,0,1,4,3,3,2,1,2,2,3], txab:[0,0,0,0,0.01,0.01,0.01,0.01,0,0.01,0.01,0.01] },
      't2': { circ:[1061,1212,1505,1448,1599,1599,1574,1681,1594,1694,1567,1518], vand:[5,13,12,6,8,6,1,5,3,4,4,10], thp:[5.73,45.93,20.4,11.65,8.82,9.22,0.65,7.82,3.87,4.77,8.83,26.18], vagab:[0,3,1,2,1,0,0,0,0,3,0,2], txab:[0,0.01,0,0.01,0,0,0,0,0,0.01,0,0.01] },
      't3': { circ:[1140,1442,1654,1566,1814,1749,1770,1898,1775,1979,1780,1753], vand:[8,6,12,6,4,2,6,5,6,1,6,5], thp:[14.3,8.47,18.35,9.03,7.78,8.6,15.08,24.55,10.9,5.03,11.97,6.22], vagab:[0,0,1,0,0,3,19,9,6,4,3,3], txab:[0,0,0,0,0,0.01,0.06,0.03,0.02,0.01,0.01,0.01] }
    },
    '2026': {
      't0': { circ:[999,1216,1585,1573,1730,1717,388,null,null,null,null,null], vand:[10,18,27,12,11,20,7,null,null,null,null,null], thp:[9.8,22.77,33.15,13.67,11.05,38.55,7.98,null,null,null,null,null], vagab:[5,9,8,3,2,2,0,null,null,null,null,null], txab:[0.02,0.03,0.02,0.01,0.01,0.01,0,null,null,null,null,null] },
      't1': { circ:[1603,1870,2472,2356,2652,2624,574,null,null,null,null,null], vand:[21,16,18,35,43,30,4,null,null,null,null,null], thp:[28.92,22.4,29.18,62.12,62.95,55.08,7.45,null,null,null,null,null], vagab:[2,0,0,2,0,0,0,null,null,null,null,null], txab:[0.01,0,0,0.01,0,0,0,null,null,null,null,null] },
      't2': { circ:[1319,1457,1571,1535,1697,1654,346,null,null,null,null,null], vand:[6,7,10,8,5,6,2,null,null,null,null,null], thp:[8.85,14.18,18.88,20.25,10.9,14.63,2,null,null,null,null,null], vagab:[0,2,0,1,0,2,0,null,null,null,null,null], txab:[0,0.01,0,0,0,0.01,0,null,null,null,null,null] },
      't3': { circ:[1550,1561,1668,1693,1798,1790,391,null,null,null,null,null], vand:[3,5,9,6,12,11,6,null,null,null,null,null], thp:[4.8,3.43,23.68,12.6,19.13,17.52,7,null,null,null,null,null], vagab:[0,1,10,4,4,3,0,null,null,null,null,null], txab:[0,0,0.03,0.01,0.01,0.01,0,null,null,null,null,null] }
    }
  }
};