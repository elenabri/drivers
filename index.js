const express = require('express');
const axios = require('axios');
const ExcelJS = require('exceljs');

const app = express();

app.use(express.json({
    limit: '10mb'
}));


// ============================================================
// MOYSKLAD
// ============================================================

const MS_TOKEN = '562d927aad09e55f49bed5ebd3751c7ccfd19a23';

if (!MS_TOKEN) {
    console.warn('WARNING: MS_TOKEN не задан');
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
// ГЛАВНАЯ СТРАНИЦА
// ============================================================

app.get('/', (req, res) => {

    res.sendFile(
        __dirname + '/index.html'
    );

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


        rows.push(
            ...batch
        );


        const total =
            response.data.meta?.size ||
            rows.length;


        console.log(
            entity,
            rows.length,
            '/',
            total
        );


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
// НОРМАЛИЗАЦИЯ ИМЕНИ
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
        '=========================================='
    );

    console.log(
        `ОТЧЁТ ПО ВОДИТЕЛЯМ: ${from} -> ${to}`
    );

    console.log(
        '=========================================='
    );


    // ========================================================
    // ЗАГРУЖАЕМ ДАННЫЕ
    // ========================================================

    const invoices =
        await getAll('invoicein');


    const demands =
        await getAll('demand');


    const counterparties =
        await getAll('counterparty');


    console.log(
        'Счетов:',
        invoices.length
    );

    console.log(
        'Отгрузок:',
        demands.length
    );

    console.log(
        'Контрагентов:',
        counterparties.length
    );


    // ========================================================
    // КОНТРАГЕНТЫ С ТЕГОМ "НАЕМНИКИ"
    // ========================================================

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


    // ========================================================
    // ПЕРИОД
    // ========================================================

    const fromDate =
        new Date(
            `${from}T00:00:00`
        );


    const toDate =
        new Date(
            `${to}T23:59:59.999`
        );


    // ========================================================
    // ОТЧЁТ
    // ========================================================

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


        // ----------------------------------------------------
        // Только выбранный период
        // ----------------------------------------------------

        if (
            invoiceDate < fromDate ||
            invoiceDate > toDate
        ) {
            continue;
        }


        // ====================================================
        // ПОЛНЫЙ СЧЁТ С КОНТРАГЕНТОМ
        // ====================================================

        let fullInvoice;


        try {

            fullInvoice = (
                await api.get(
                    `/entity/invoicein/${invoice.id}?expand=agent`
                )
            ).data;

        } catch (error) {

            console.error(
                'Ошибка получения счета:',
                invoice.id,
                error.response?.data ||
                error.message
            );

            continue;

        }


        if (!fullInvoice.agent?.name) {
            continue;
        }


        const counterparty =
            fullInvoice.agent.name;


        const normalizedCounterparty =
            normalizeName(counterparty);


        // ====================================================
        // ТОЛЬКО "НАЕМНИКИ"
        // ====================================================

        if (
            !mercenaryNames.has(
                normalizedCounterparty
            )
        ) {

            continue;

        }


        const day =
            fullInvoice.moment.substring(
                0,
                10
            );


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
                demand.moment.substring(
                    0,
                    10
                );


            // ------------------------------------------------
            // Отгрузка в тот же день
            // ------------------------------------------------

            if (demandDay !== day) {
                continue;
            }


            // =================================================
            // ВОДИТЕЛЬ
            // =================================================

            const driverAttr =
                demand.attributes
                    ?.find(
                        a =>
                            a.name === 'Водитель'
                    )
                    ?.value?.name || '';


            if (!driverAttr) {
                continue;
            }


            const normalizedDriver =
                normalizeName(
                    driverAttr
                );


            // =================================================
            // КОНТРАГЕНТ ОТГРУЗКИ
            // =================================================

            const demandCounterparty =
                demand.agent?.name || '';


            // =================================================
            // СОПОСТАВЛЕНИЕ
            // =================================================

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


            // =================================================
            // СУММА ОТГРУЗКИ
            // =================================================

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


        // ====================================================
        // ЕСЛИ ОТГРУЗОК НЕТ
        // ====================================================

        if (
            matchedDemands.length === 0
        ) {

            continue;

        }


        // ====================================================
        // ДОБАВЛЯЕМ В ОТЧЁТ
        // ====================================================

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
                            .map(
                                d => d.driver
                            )
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

    report.sort(
        (a, b) => {

            if (
                a.date !== b.date
            ) {

                return a.date.localeCompare(
                    b.date
                );

            }


            return a.counterparty.localeCompare(
                b.counterparty
            );

        }
    );


    console.log(
        'Итого строк:',
        report.length
    );


    return report;

}


// ============================================================
// API: ПОЛУЧИТЬ ОТЧЁТ
//
// GET
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


            // ------------------------------------------------
            // Проверяем даты
            // ------------------------------------------------

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


            // ------------------------------------------------
            // Создаём отчёт
            // ------------------------------------------------

            const report =
                await buildDriverReport(
                    from,
                    to
                );


            // ------------------------------------------------
            // Возвращаем JSON
            // ------------------------------------------------

            res.json({

                success: true,

                from,

                to,

                count:
                    report.length,

                rows:
                    report

            });


        } catch (error) {

            console.error(
                'DRIVER REPORT ERROR:',
                error.response?.data ||
                error.message ||
                error
            );


            res.status(500).json({

                success: false,

                error:
                    error.response?.data ||
                    error.message ||
                    'Ошибка получения отчёта'

            });

        }

    }
);


// ============================================================
// API: СОЗДАТЬ EXCEL
//
// ВАЖНО:
//
// Этот endpoint НЕ обращается к МойСклад.
//
// Он получает уже готовые строки от браузера.
//
// POST
// /api/driver-report.xlsx
// ============================================================

app.post(
    '/api/driver-report.xlsx',
    async (req, res) => {

        try {

            const {
                from,
                to,
                rows
            } = req.body;


            // ------------------------------------------------
            // Проверяем данные
            // ------------------------------------------------

            if (!from || !to) {

                return res.status(400).json({

                    error:
                        'Не указан период'

                });

            }


            if (!Array.isArray(rows)) {

                return res.status(400).json({

                    error:
                        'Не переданы строки отчёта'

                });

            }


            console.log(
                `Создание Excel: ${rows.length} строк`
            );


            // =================================================
            // СОЗДАЁМ EXCEL
            // =================================================

            const workbook =
                new ExcelJS.Workbook();


            const sheet =
                workbook.addWorksheet(
                    'Отчет'
                );


            // =================================================
            // КОЛОНКИ
            // =================================================

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


            // =================================================
            // СТРОКИ
            // =================================================

            for (const row of rows) {

                const demandsText =
                    (row.demands || [])
                        .map(
                            demand =>
                                `${demand.number} (${Number(demand.sum || 0).toFixed(2)})`
                        )
                        .join(', ');


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
                        Number(
                            row.invoiceSum || 0
                        ),

                    shipmentSum:
                        Number(
                            row.shipmentSum || 0
                        ),

                    difference:
                        Number(
                            row.difference || 0
                        ),

                    demands:
                        demandsText

                });

            }


            // =================================================
            // ЗАКРЕПИТЬ ПЕРВУЮ СТРОКУ
            // =================================================

            sheet.views = [

                {
                    state: 'frozen',

                    ySplit: 1
                }

            ];


            // =================================================
            // ФИЛЬТР
            // =================================================

            sheet.autoFilter = {

                from: 'A1',

                to: 'H1'

            };


            // =================================================
            // ФОРМАТ ЧИСЕЛ
            // =================================================

            for (
                let i = 2;
                i <= sheet.rowCount;
                i++
            ) {

                sheet.getCell(
                    `E${i}`
                ).numFmt =
                    '#,##0.00';


                sheet.getCell(
                    `F${i}`
                ).numFmt =
                    '#,##0.00';


                sheet.getCell(
                    `G${i}`
                ).numFmt =
                    '#,##0.00';


                const difference =
                    Number(
                        sheet.getCell(
                            `G${i}`
                        ).value || 0
                    );


                // --------------------------------------------
                // Положительная разница
                // --------------------------------------------

                if (
                    difference > 0
                ) {

                    sheet.getCell(
                        `G${i}`
                    ).fill = {

                        type: 'pattern',

                        pattern: 'solid',

                        fgColor: {

                            argb:
                                'FFC6EFCE'

                        }

                    };

                }


                // --------------------------------------------
                // Отрицательная разница
                // --------------------------------------------

                if (
                    difference < 0
                ) {

                    sheet.getCell(
                        `G${i}`
                    ).fill = {

                        type: 'pattern',

                        pattern: 'solid',

                        fgColor: {

                            argb:
                                'FFFFC7CE'

                        }

                    };

                }

            }


            // =================================================
            // EXCEL В ПАМЯТЬ
            // =================================================

            const buffer =
                await workbook.xlsx.writeBuffer();


            // =================================================
            // ИМЯ ФАЙЛА
            // =================================================

            const fileName =
                `driver-report-${from}-${to}.xlsx`;


            // =================================================
            // ОТДАЁМ ФАЙЛ
            // =================================================

            res.setHeader(
                'Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            );


            res.setHeader(
                'Content-Disposition',
                `attachment; filename="${fileName}"`
            );


            res.send(buffer);


        } catch (error) {

            console.error(
                'EXCEL ERROR:',
                error.response?.data ||
                error.message ||
                error
            );


            res.status(500).json({

                error:
                    error.message ||
                    'Ошибка создания Excel'

            });

        }

    }
);


// ============================================================
// VERCEL
// ============================================================

module.exports = app;
