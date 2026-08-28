const express = require('express');
const axios = require('axios');
const ExcelJS = require('exceljs');

const app = express();


// ============================================================
// MOYSKLAD
// ============================================================

const MS_TOKEN = '562d927aad09e55f49bed5ebd3751c7ccfd19a23';

if (!MS_TOKEN) {
    console.warn('MS_TOKEN не задан');
}

const api = axios.create({
    baseURL: 'https://api.moysklad.ru/api/remap/1.2',
    headers: {
        Authorization: `Bearer ${MS_TOKEN}`,
        Accept: 'application/json;charset=utf-8'
    },
    timeout: 30000
});


// ============================================================
// HTML
// ============================================================

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});


// ============================================================
// ПОЛУЧИТЬ ВСЕ ЗАПИСИ
// ============================================================

async function getAll(entity) {

    let offset = 0;
    const limit = 1000;

    const rows = [];

    while (true) {

        const response = await api.get(
            `/entity/${entity}?limit=${limit}&offset=${offset}`
        );

        const batch =
            response.data.rows || [];

        rows.push(...batch);

        const total =
            response.data.meta?.size ||
            rows.length;

        if (rows.length >= total) {
            break;
        }

        if (!batch.length) {
            break;
        }

        offset += limit;
    }

    return rows;
}


// ============================================================
// НОРМАЛИЗАЦИЯ
// ============================================================

function normalizeName(name) {

    return String(name || '')
        .trim()
        .replace(/\s+/g, ' ')
        .toLowerCase();
}


// ============================================================
// СОЗДАНИЕ ОТЧЁТА
// ============================================================

async function buildDriverReport(from, to) {

    console.log(
        `Создание отчёта: ${from} -> ${to}`
    );


    // --------------------------------------------------------
    // Загружаем данные
    // --------------------------------------------------------

    const [
        invoices,
        demands,
        counterparties
    ] = await Promise.all([

        getAll('invoicein'),

        getAll('demand'),

        getAll('counterparty')

    ]);


    console.log(
        'Счета:',
        invoices.length
    );

    console.log(
        'Отгрузки:',
        demands.length
    );

    console.log(
        'Контрагенты:',
        counterparties.length
    );


    // --------------------------------------------------------
    // Контрагенты "наемники"
    // --------------------------------------------------------

    const mercenaryNames =
        new Set();


    for (const cp of counterparties) {

        if (
            cp.tags &&
            cp.tags.includes('наемники')
        ) {

            mercenaryNames.add(
                normalizeName(cp.name)
            );
        }
    }


    console.log(
        'Наемников:',
        mercenaryNames.size
    );


    // --------------------------------------------------------
    // Период
    // --------------------------------------------------------

    const fromDate =
        new Date(`${from}T00:00:00`);

    const toDate =
        new Date(`${to}T23:59:59.999`);


    const report = [];


    // ========================================================
    // ОБРАБОТКА СЧЕТОВ
    // ========================================================

    for (const invoice of invoices) {

        if (!invoice.moment) {
            continue;
        }


        const invoiceDate =
            new Date(invoice.moment);


        // Счёт вне периода
        if (
            invoiceDate < fromDate ||
            invoiceDate > toDate
        ) {
            continue;
        }


        // ----------------------------------------------------
        // Полный счёт
        // ----------------------------------------------------

        const fullInvoice = (
            await api.get(
                `/entity/invoicein/${invoice.id}?expand=agent`
            )
        ).data;


        if (!fullInvoice.agent?.name) {
            continue;
        }


        const counterparty =
            fullInvoice.agent.name;


        const normalizedCounterparty =
            normalizeName(counterparty);


        // ----------------------------------------------------
        // Только наемники
        // ----------------------------------------------------

        if (
            !mercenaryNames.has(
                normalizedCounterparty
            )
        ) {
            continue;
        }


        const day =
            fullInvoice.moment.substring(0, 10);


        const matchedDemands = [];

        let shipmentSum = 0;


        // ====================================================
        // ПОИСК ОТГРУЗОК
        // ====================================================

        for (const demand of demands) {

            if (!demand.moment) {
                continue;
            }


            const demandDay =
                demand.moment.substring(0, 10);


            // Отгрузка должна быть в тот же день
            if (demandDay !== day) {
                continue;
            }


            // ------------------------------------------------
            // Водитель
            // ------------------------------------------------

            const driverAttr =
                demand.attributes
                    ?.find(
                        a => a.name === 'Водитель'
                    )
                    ?.value?.name || '';


            if (!driverAttr) {
                continue;
            }


            const normalizedDriver =
                normalizeName(driverAttr);


            // ------------------------------------------------
            // Контрагент отгрузки
            // ------------------------------------------------

            const demandCounterparty =
                demand.agent?.name || '';


            // ------------------------------------------------
            // Сопоставление
            // ------------------------------------------------

            const driverMatches =
                normalizedCounterparty.includes(
                    normalizedDriver
                );


            const counterpartyMatches =
                normalizeName(
                    demandCounterparty
                ) === normalizedCounterparty;


            if (
                !driverMatches &&
                !counterpartyMatches
            ) {
                continue;
            }


            // ------------------------------------------------
            // Сумма
            // ------------------------------------------------

            if (
                (demand.sum || 0) <= 0
            ) {
                continue;
            }


            const sum =
                demand.sum / 100;


            shipmentSum += sum;


            matchedDemands.push({

                id:
                    demand.id,

                number:
                    demand.name,

                sum,

                driver:
                    driverAttr,

                counterparty:
                    demandCounterparty

            });

        }


        // ----------------------------------------------------
        // Нет отгрузок
        // ----------------------------------------------------

        if (!matchedDemands.length) {
            continue;
        }


        // ----------------------------------------------------
        // Добавляем строку
        // ----------------------------------------------------

        report.push({

            invoiceId:
                fullInvoice.id,

            invoice:
                fullInvoice.name,

            date:
                day,

            counterparty,

            driver:
                [
                    ...new Set(
                        matchedDemands
                            .map(d => d.driver)
                            .filter(Boolean)
                    )
                ].join(', '),

            invoiceSum:
                fullInvoice.sum / 100,

            shipmentSum,

            difference:
                shipmentSum -
                fullInvoice.sum / 100,

            demands:
                matchedDemands

        });

    }


    // ========================================================
    // СОРТИРОВКА
    // ========================================================

    report.sort((a, b) => {

        if (a.date !== b.date) {

            return a.date.localeCompare(
                b.date
            );
        }

        return a.counterparty.localeCompare(
            b.counterparty
        );

    });


    console.log(
        'Итого строк:',
        report.length
    );


    return report;
}


// ============================================================
// JSON ОТЧЁТ
//
// /api/driver-report?from=2026-08-01&to=2026-08-28
// ============================================================

app.get(
    '/api/driver-report',
    async (req, res) => {

        try {

            const {
                from,
                to
            } = req.query;


            if (!from || !to) {

                return res.status(400).json({

                    success: false,

                    error:
                        'Необходимо указать from и to'

                });
            }


            if (from > to) {

                return res.status(400).json({

                    success: false,

                    error:
                        'Дата начала больше даты окончания'

                });
            }


            const report =
                await buildDriverReport(
                    from,
                    to
                );


            res.json({

                success: true,

                from,

                to,

                count:
                    report.length,

                rows:
                    report

            });


        } catch (e) {

            console.error(
                'DRIVER REPORT ERROR:',
                e.response?.data ||
                e.message ||
                e
            );


            res.status(500).json({

                success: false,

                error:
                    e.response?.data ||
                    e.message ||
                    'Ошибка получения отчёта'

            });

        }

    }
);


// ============================================================
// EXCEL
//
// /api/driver-report.xlsx?from=2026-08-01&to=2026-08-28
// ============================================================

app.get(
    '/api/driver-report.xlsx',
    async (req, res) => {

        try {

            const {
                from,
                to
            } = req.query;


            if (!from || !to) {

                return res.status(400).json({

                    error:
                        'Необходимо указать from и to'

                });
            }


            if (from > to) {

                return res.status(400).json({

                    error:
                        'Дата начала больше даты окончания'

                });
            }


            const report =
                await buildDriverReport(
                    from,
                    to
                );


            // ------------------------------------------------
            // Создаём Excel
            // ------------------------------------------------

            const workbook =
                new ExcelJS.Workbook();


            const sheet =
                workbook.addWorksheet(
                    'Отчет'
                );


            sheet.columns = [

                {
                    header: 'Дата',
                    key: 'date',
                    width: 14
                },

                {
                    header: 'Контрагент',
                    key: 'counterparty',
                    width: 50
                },

                {
                    header: 'Водитель',
                    key: 'driver',
                    width: 30
                },

                {
                    header: 'Счет',
                    key: 'invoice',
                    width: 15
                },

                {
                    header: 'Сумма счета',
                    key: 'invoiceSum',
                    width: 18
                },

                {
                    header: 'Сумма отгрузок',
                    key: 'shipmentSum',
                    width: 20
                },

                {
                    header: 'Разница',
                    key: 'difference',
                    width: 18
                },

                {
                    header: 'Отгрузки',
                    key: 'demands',
                    width: 45
                }

            ];


            // ------------------------------------------------
            // Заполняем Excel
            // ------------------------------------------------

            for (const row of report) {

                sheet.addRow({

                    date:
                        row.date,

                    counterparty:
                        row.counterparty,

                    driver:
                        row.driver,

                    invoice:
                        row.invoice,

                    invoiceSum:
                        row.invoiceSum,

                    shipmentSum:
                        row.shipmentSum,

                    difference:
                        row.difference,

                    demands:
                        (row.demands || [])
                            .map(
                                d =>
                                    `${d.number} (${Number(d.sum).toFixed(2)})`
                            )
                            .join(', ')

                });

            }


            // ------------------------------------------------
            // Закрепляем первую строку
            // ------------------------------------------------

            sheet.views = [

                {
                    state: 'frozen',
                    ySplit: 1
                }

            ];


            // ------------------------------------------------
            // Фильтр
            // ------------------------------------------------

            sheet.autoFilter = {

                from: 'A1',

                to: 'H1'

            };


            // ------------------------------------------------
            // Формат денег
            // ------------------------------------------------

            for (
                let i = 2;
                i <= sheet.rowCount;
                i++
            ) {

                sheet.getCell(`E${i}`)
                    .numFmt =
                    '#,##0.00';

                sheet.getCell(`F${i}`)
                    .numFmt =
                    '#,##0.00';

                sheet.getCell(`G${i}`)
                    .numFmt =
                    '#,##0.00';


                const diff =
                    Number(
                        sheet.getCell(
                            `G${i}`
                        ).value
                    );


                if (diff > 0) {

                    sheet.getCell(
                        `G${i}`
                    ).fill = {

                        type: 'pattern',

                        pattern: 'solid',

                        fgColor: {
                            argb: 'FFC6EFCE'
                        }

                    };

                }


                if (diff < 0) {

                    sheet.getCell(
                        `G${i}`
                    ).fill = {

                        type: 'pattern',

                        pattern: 'solid',

                        fgColor: {
                            argb: 'FFFFC7CE'
                        }

                    };

                }

            }


            // ------------------------------------------------
            // Excel в память
            // ------------------------------------------------

            const buffer =
                await workbook.xlsx.writeBuffer();


            const fileName =
                `driver-report-${from}-${to}.xlsx`;


            res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            );


            res.setHeader(
                'Content-Disposition',
                `attachment; filename="${fileName}"`
            );


            res.send(buffer);


        } catch (e) {

            console.error(
                'EXCEL ERROR:',
                e.response?.data ||
                e.message ||
                e
            );


            res.status(500).json({

                error:
                    e.response?.data ||
                    e.message ||
                    'Ошибка создания Excel'

            });

        }

    }
);


// ============================================================
// VERCEL
// ============================================================

module.exports = app;
