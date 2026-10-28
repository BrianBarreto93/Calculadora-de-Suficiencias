/*
 * Datos de ejemplo: los mismos de las hojas IND01_Churn e IND02_Instalaciones
 * de Plantilla_Suficiencia.xlsx.
 */
(function (raiz, fabrica) {
  if (typeof module === 'object' && module.exports) module.exports = fabrica();
  else raiz.SuficienciaEjemplos = fabrica();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  return function crearEjemplos() {
    return [
      {
        id: 'ej-churn',
        nombre: 'Churn mensual pospago',
        sentido: 'baja',
        unidad: '%',
        ejecutado: 2.8,
        meta: 2,
        ideal: null,
        totalCasos: 12000,
        dolores: [
          { id: 'd1', nombre: 'Precio / oferta de la competencia', casos: 6000, atribucionManual: null, evidencia: 'Encuesta de salida Q2-2026 (n = 12.000 bajas)' },
          { id: 'd2', nombre: 'Fallas de red / cobertura', casos: 3600, atribucionManual: null, evidencia: 'Cruce de bajas con mapa de caídas de red (últimos 90 días)' },
          { id: 'd3', nombre: 'Reclamos no resueltos', casos: 3000, atribucionManual: null, evidencia: 'CRM: reclamo abierto más de 7 días antes de la baja' },
          { id: 'd4', nombre: 'Errores de facturación', casos: 1800, atribucionManual: null, evidencia: 'Notas de crédito emitidas en los 3 meses previos a la baja' },
          { id: 'd5', nombre: 'No gestionable (mudanza, fallecimiento)', casos: 1200, atribucionManual: null, evidencia: 'Motivo de baja declarado en tienda' }
        ],
        iniciativas: [
          { id: 'i1', nombre: 'Nuevo portafolio de planes competitivos', efectividad: { d1: 0.45 }, ejecucion: 0.4, estado: 'Retrasada', responsable: 'Marketing', fecha: '2027-03-31', evidencia: 'Piloto: 45 % de los clientes que se iban por precio aceptó el nuevo plan' },
          { id: 'i2', nombre: 'Ofertas de retención con modelo predictivo', efectividad: { d1: 0.1, d3: 0.15 }, ejecucion: 1, estado: 'En plan', responsable: 'Retención', fecha: '2026-12-15', evidencia: 'Piloto Q1-2026: retuvo 10 % (precio) y 15 % (reclamos) de clientes en riesgo' },
          { id: 'i3', nombre: 'Densificación de red en zonas críticas', efectividad: { d2: 0.3 }, ejecucion: 0.6, estado: 'En riesgo', responsable: 'Red', fecha: '2027-06-30', evidencia: 'Simulación: 30 % de las quejas de red están en sitios planificados' },
          { id: 'i4', nombre: 'Resolución en el primer contacto', efectividad: { d3: 0.3, d4: 0.25 }, ejecucion: 0.9, estado: 'En plan', responsable: 'Experiencia cliente', fecha: '2027-01-31', evidencia: 'Piloto zona sur: −30 % reclamos reiterados y −25 % errores de factura' },
          { id: 'i5', nombre: 'Campaña de imagen de marca en redes', efectividad: { d1: 0.01 }, ejecucion: 1, estado: 'En plan', responsable: 'Comunicaciones', fecha: '2026-11-30', evidencia: '' }
        ]
      },
      {
        id: 'ej-instalaciones',
        nombre: 'Instalaciones en fecha (FTTH)',
        sentido: 'sube',
        unidad: '%',
        ejecutado: 82,
        meta: 92,
        ideal: null,
        totalCasos: 9000,
        dolores: [
          { id: 'd1', nombre: 'Falta de técnicos en la zona', casos: 4050, atribucionManual: null, evidencia: 'Órdenes reprogramadas por falta de cuadrilla (sistema de despacho)' },
          { id: 'd2', nombre: 'Cliente ausente o dirección errónea', casos: 2700, atribucionManual: null, evidencia: 'Visitas fallidas registradas por el técnico' },
          { id: 'd3', nombre: 'Falta de equipos (ONT / módem) en almacén', casos: 1800, atribucionManual: null, evidencia: 'Órdenes en espera por stock (ERP de logística)' },
          { id: 'd4', nombre: 'Factibilidad mal validada en la venta', casos: 1350, atribucionManual: null, evidencia: 'Órdenes rechazadas por falta de puerto o red' }
        ],
        iniciativas: [
          { id: 'i1', nombre: 'Agendamiento con confirmación por WhatsApp', efectividad: { d2: 0.7 }, ejecucion: 1, estado: 'En plan', responsable: 'Experiencia cliente', fecha: '2026-12-31', evidencia: 'Piloto zona norte: −70 % de visitas fallidas' },
          { id: 'i2', nombre: 'Contratación de técnicos tercerizados', efectividad: { d1: 0.6 }, ejecucion: 0.9, estado: 'En riesgo', responsable: 'Operaciones', fecha: '2027-02-28', evidencia: 'Plan de capacidad: +60 % de cuadrillas en zonas con atraso' },
          { id: 'i3', nombre: 'Reposición automática de inventario', efectividad: { d3: 0.8, d1: 0.1 }, ejecucion: 1, estado: 'En plan', responsable: 'Logística', fecha: '2027-01-15', evidencia: 'Modelo de stock mínimo: cubre 80 % de los quiebres de stock' },
          { id: 'i4', nombre: 'Validación de factibilidad en el punto de venta', efectividad: { d4: 0.7, d2: 0.1 }, ejecucion: 1, estado: 'En plan', responsable: 'Ventas', fecha: '2027-03-31', evidencia: 'App de factibilidad: 70 % de los rechazos se detectan antes de vender' }
        ]
      }
    ];
  };
});
