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
        Accept: 'application/json;charset=utf-8',
        'Accept-Encoding': 'gzip'
    },

    timeout: 60000
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
// ПОЛУЧЕНИЕ ВСЕХ ЗАПИСЕЙ
// ============================================================

async function getAll(entity) {

    let offset = 0;

    const limit = 1000;

    const rows = [];

    while (true) {

        const response = await api.get(
            `/entity/${entity}`,
            {
                params: {
                    limit,
                    offset
                }
            }
        );

        const batch =
            response.data?.rows || [];

        rows.push(
            ...batch
        );

        const total =
            Number(
                response.data?.meta?.size ||
                0
            );

        console.log(
            `${entity}: ${rows.length}${total ? ` / ${total}` : ''}`
        );

        if (
            batch.length === 0 ||
            rows.length >= total
        ) {
            break;
        }

        offset += limit;
    }

    return rows;
}


// ============================================================
// НОРМАЛИЗАЦИЯ СТРОКИ
// ============================================================

function normalizeName(value) {

    return String(
        value || ''
    )
        .trim()
        .replace(/\s+/g, ' ')
        .toLowerCase();
}


// ============================================================
// ПОЛУЧЕНИЕ АТРИБУТА
// ============================================================

function getAttribute(entity, attributeName) {

    const attribute =
        entity?.attributes?.find(
            item =>
                normalizeName(item.name) ===
                normalizeName(attributeName)
        );

    if (!attribute) {
        return '';
    }

    if (
        attribute.value &&
        typeof attribute.value === 'object'
    ) {
        return (
            attribute.value.name ||
            attribute.value.value ||
            ''
        );
    }

    return String(
        attribute.value ?? ''
    );
}


// ============================================================
// ПОЛУЧЕНИЕ ВОДИТЕЛЯ
// ============================================================

function getDriver(demand) {

    return (
        getAttribute(
            demand,
            'Водитель'
        ) ||
        getAttribute(
            demand,
            'водитель'
        ) ||
        ''
    );
}


// ============================================================
// ПОЛУЧЕНИЕ СУММЫ
// ============================================================

function moneyFromMoySklad(value) {

    const number =
        Number(value || 0);

    return number / 100;
}


// ============================================================
// ПОЛУЧЕНИЕ ОТГРУЗОК
// ============================================================

async function getDemandsForPeriod(
    demands,
    from,
    to
) {

    const result = [];

    for (
        const demand
        of demands
    ) {

        if (!demand.moment) {
            continue;
        }

        const day =
            demand.moment.substring(
                0,
                10
            );

        if (
            day < from ||
            day > to
        ) {
            continue;
        }

        const driver =
            getDriver(demand);

        if (!driver) {
            continue;
        }

        const sum =
            moneyFromMoySklad(
                demand.sum
            );

        if (
            !Number.isFinite(sum) ||
            sum <= 0
        ) {
            continue;
        }

        result.push({

            id:
                demand.id,

            number:
                demand.name || '',

            date:
                day,

            driver:
                driver,

            counterparty:
                demand.agent?.name || '',

            sum:
                sum

        });
    }

    return result;
}


// ============================================================
// ПОЛУЧЕНИЕ СТАТУСА СЧЁТА
// ============================================================

function getInvoiceStatus(invoice) {

    return (
        invoice?.state?.name ||
        'Без статуса'
    );
}


// ============================================================
// ПОПЫТКА ОПРЕДЕЛИТЬ КОНТРАГЕНТА
// ============================================================

function getInvoiceCounterparty(invoice) {

    return (
        invoice?.agent?.name ||
        ''
    );
}


// ============================================================
// СОПОСТАВЛЕНИЕ ОТГРУЗКИ СО СЧЁТОМ
// ============================================================

function demandBelongsToInvoice(
    demand,
    invoiceCounterparty,
    driver
) {

    const demandCounterparty =
        normalizeName(
            demand.counterparty
        );

    const invoiceCp =
        normalizeName(
            invoiceCounterparty
        );

    const normalizedDriver =
        normalizeName(
            driver
        );

    // --------------------------------------------------------
    // Основное совпадение:
    // контрагент счета = контрагент отгрузки
    // --------------------------------------------------------

    if (
        demandCounterparty &&
        invoiceCp &&
        demandCounterparty === invoiceCp
    ) {
        return true;
    }

    // --------------------------------------------------------
    // Дополнительная проверка через водителя
    // --------------------------------------------------------

    if (
        normalizedDriver &&
        (
            demandCounterparty.includes(
                normalizedDriver
            ) ||
            normalizedDriver.includes(
                demandCounterparty
            )
        )
    ) {
        return true;
    }

    return false;
}


// ============================================================
// ОСНОВНОЙ ОТЧЁТ
// ============================================================

async function buildDriverReport(
    from,
    to
) {

    console.log(
        '=========================================='
    );

    console.log(
        `ОТЧЁТ: ${from} -> ${to}`
    );

    console.log(
        '=========================================='
    );


    // ========================================================
    // Загружаем основные сущности
    // ========================================================

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
    // НАЕМНИКИ
    // ========================================================

    const mercenaryNames =
        new Set();

    for (
        const cp
        of counterparties
    ) {

        const tags =
            Array.isArray(cp.tags)
                ? cp.tags
                : [];

        const hasMercenaryTag =
            tags.some(
                tag =>
                    normalizeName(tag) ===
                    'наемники'
            );

        if (
            hasMercenaryTag
        ) {

            mercenaryNames.add(
                normalizeName(
                    cp.name
                )
            );

        }
    }


    console.log(
        'Наемников:',
        mercenaryNames.size
    );


    // ========================================================
    // Отгрузки выбранного периода
    // ========================================================

    const periodDemands =
        await getDemandsForPeriod(
            demands,
            from,
            to
        );


    console.log(
        'Отгрузок за период:',
        periodDemands.length
    );


    const report = [];


    // ========================================================
    // ОБРАБАТЫВАЕМ СЧЕТА
    // ========================================================

    for (
        const invoice
        of invoices
    ) {

        if (!invoice.moment) {
            continue;
        }


        const day =
            invoice.moment.substring(
                0,
                10
            );


        // ----------------------------------------------------
        // Период
        // ----------------------------------------------------

        if (
            day < from ||
            day > to
        ) {
            continue;
        }


        // ----------------------------------------------------
        // Получаем полный счет.
        //
        // ВАЖНО:
        // expand=agent,state позволяет сразу получить
        // контрагента и статус.
        // ----------------------------------------------------

        let fullInvoice;

        try {

            fullInvoice =
                (
                    await api.get(
                        `/entity/invoicein/${invoice.id}`,
                        {
                            params: {
                                expand:
                                    'agent,state'
                            }
                        }
                    )
                ).data;

        } catch (error) {

            console.error(
                'Ошибка получения счета',
                invoice.id,
                error.response?.data ||
                error.message
            );

            continue;
        }


        // ----------------------------------------------------
        // Контрагент
        // ----------------------------------------------------

        const counterparty =
            getInvoiceCounterparty(
                fullInvoice
            );

        if (!counterparty) {
            continue;
        }


        // ----------------------------------------------------
        // Только "Наемники"
        // ----------------------------------------------------

        if (
            !mercenaryNames.has(
                normalizeName(
                    counterparty
                )
            )
        ) {
            continue;
        }


        // ----------------------------------------------------
        // Статус
        // ----------------------------------------------------

        const status =
            getInvoiceStatus(
                fullInvoice
            );


        // ----------------------------------------------------
        // Сумма счета
        // ----------------------------------------------------

        const invoiceSum =
            moneyFromMoySklad(
                fullInvoice.sum
            );


        // ====================================================
        // Ищем отгрузки этого контрагента в этот день
        // ====================================================

        const matchedDemands =
            periodDemands.filter(
                demand => {

                    if (
                        demand.date !== day
                    ) {
                        return false;
                    }

                    return demandBelongsToInvoice(
                        demand,
                        counterparty,
                        demand.driver
                    );
                }
            );


        if (
            matchedDemands.length === 0
        ) {
            continue;
        }


        // ====================================================
        // СУММА ОТГРУЗОК
        // ====================================================

        const shipmentSum =
            matchedDemands.reduce(
                (
                    total,
                    demand
                ) =>
                    total +
                    Number(
                        demand.sum || 0
                    ),
                0
            );


        // ====================================================
        // 20%
        // ====================================================

        const twentyPercent =
            shipmentSum * 0.20;


        // ====================================================
        // РАЗНИЦА
        //
        // 100000 - 80000 - 16000
        // = 4000
        // ====================================================

        const difference =
            invoiceSum -
            shipmentSum -
            twentyPercent;


        // ====================================================
        // ВОДИТЕЛИ
        // ====================================================

        const drivers =
            [
                ...new Set(
                    matchedDemands
                        .map(
                            demand =>
                                demand.driver
                        )
                        .filter(Boolean)
                )
            ];


        // ====================================================
        // ДОБАВЛЯЕМ СТРОКУ
        // ====================================================

        report.push({

            invoiceId:
                fullInvoice.id,

            invoice:
                fullInvoice.name || '',

            status:
                status,

            date:
                day,

            counterparty:
                counterparty,

            driver:
                drivers.join(', '),

            invoiceSum:
                invoiceSum,

            shipmentSum:
                shipmentSum,

            twentyPercent:
                twentyPercent,

            difference:
                difference,

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
                b.counterparty,
                'ru'
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
// ============================================================

app.get(
    '/api/driver-report',
    async (req, res) => {

        try {

            const {
                from,
                to
            } = req.query;


            if (
                !from ||
                !to
            ) {

                return res
                    .status(400)
                    .json({

                        success:
                            false,

                        error:
                            'Необходимо указать from и to'

                    });
            }


            if (
                from > to
            ) {

                return res
                    .status(400)
                    .json({

                        success:
                            false,

                        error:
                            'Дата начала не может быть позже даты окончания'

                    });
            }


            const report =
                await buildDriverReport(
                    from,
                    to
                );


            res.json({

                success:
                    true,

                from:
                    from,

                to:
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


            res
                .status(500)
                .json({

                    success:
                        false,

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
// Здесь НЕ обращаемся к МойСклад.
//
// Браузер присылает уже отфильтрованный список.
//
// Поэтому ошибка 1073 при скачивании не должна возникать
// из-за повторных запросов к API МойСклад.
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


            if (
                !from ||
                !to
            ) {

                return res
                    .status(400)
                    .json({

                        error:
                            'Не указан период'

                    });
            }


            if (
                !Array.isArray(rows)
            ) {

                return res
                    .status(400)
                    .json({

                        error:
                            'Не переданы строки отчёта'

                    });
            }


            if (
                rows.length === 0
            ) {

                return res
                    .status(400)
                    .json({

                        error:
                            'Нет данных для скачивания'

                    });
            }


            console.log(
                'Excel:',
                rows.length,
                'строк'
            );


            // =================================================
            // WORKBOOK
            // =================================================

            const workbook =
                new ExcelJS.Workbook();


            workbook.creator =
                'Отчёт по водителям';


            workbook.created =
                new Date();


            // =================================================
            // ЛИСТ
            // =================================================

            const sheet =
                workbook.addWorksheet(
                    'Отчет'
                );


            // =================================================
            // КОЛОНКИ
            //
            // СТАТУСА НЕТ.
            // =================================================

            sheet.columns = [

                {
                    header:
                        'Дата',

                    key:
                        'date',

                    width:
                        8

                },

                {
                    header:
                        'Контрагент',

                    key:
                        'counterparty',

                    width:
                        18

                },

                {
                    header:
                        'Водитель',

                    key:
                        'driver',

                    width:
                        14

                },

                {
                    header:
                        'Сумма счета',

                    key:
                        'invoiceSum',

                    width:
                        11

                },

                {
                    header:
                        'Сумма отгрузок',

                    key:
                        'shipmentSum',

                    width:
                        12

                },

                {
                    header:
                        '20%',

                    key:
                        'twentyPercent',

                    width:
                        8

                },

                {
                    header:
                        'Разница',

                    key:
                        'difference',

                    width:
                        11

                }

            ];


            // =================================================
            // ДАННЫЕ
            // =================================================

            for (
                const row
                of rows
            ) {

                const invoiceSum =
                    Number(
                        row.invoiceSum || 0
                    );


                const shipmentSum =
                    Number(
                        row.shipmentSum || 0
                    );


                // ---------------------------------------------
                // Всегда пересчитываем 20%
                // ---------------------------------------------

                const twentyPercent =
                    shipmentSum * 0.20;


                // ---------------------------------------------
                // Всегда пересчитываем разницу
                // ---------------------------------------------

                const difference =
                    invoiceSum -
                    shipmentSum -
                    twentyPercent;


                sheet.addRow({

                    date:
                        row.date || '',

                    counterparty:
                        row.counterparty || '',

                    driver:
                        row.driver || '',

                    invoiceSum:
                        invoiceSum,

                    shipmentSum:
                        shipmentSum,

                    twentyPercent:
                        twentyPercent,

                    difference:
                        difference

                });
            }


            // =================================================
            // ИТОГО
            // =================================================

            const dataLastRow =
                sheet.rowCount;


            const totalRow =
                sheet.addRow({

                    date:
                        '',

                    counterparty:
                        'ИТОГО',

                    driver:
                        '',

                    invoiceSum:
                        {
                            formula:
                                `SUM(D2:D${dataLastRow})`
                        },

                    shipmentSum:
                        {
                            formula:
                                `SUM(E2:E${dataLastRow})`
                        },

                    twentyPercent:
                        {
                            formula:
                                `SUM(F2:F${dataLastRow})`
                        },

                    difference:
                        {
                            formula:
                                `SUM(G2:G${dataLastRow})`
                        }

                });


            // =================================================
            // ЗАГОЛОВОК
            // =================================================

            const headerRow =
                sheet.getRow(1);


            headerRow.font = {
                bold: true
            };


            headerRow.alignment = {

                vertical:
                    'middle',

                horizontal:
                    'center',

                wrapText:
                    true

            };


            headerRow.height =
                28;


            // =================================================
            // ФОРМАТ ДЕНЕГ
            // =================================================

            for (
                let rowNumber = 2;
                rowNumber <= sheet.rowCount;
                rowNumber++
            ) {

                sheet.getCell(
                    `D${rowNumber}`
                ).numFmt =
                    '#,##0.00';


                sheet.getCell(
                    `E${rowNumber}`
                ).numFmt =
                    '#,##0.00';


                sheet.getCell(
                    `F${rowNumber}`
                ).numFmt =
                    '#,##0.00';


                sheet.getCell(
                    `G${rowNumber}`
                ).numFmt =
                    '#,##0.00';
            }


            // =================================================
            // ИТОГО
            // =================================================

            totalRow.font = {
                bold: true
            };


            totalRow.height =
                24;


            for (
                let col = 1;
                col <= 7;
                col++
            ) {

                totalRow.getCell(
                    col
                ).border = {

                    top: {
                        style:
                            'thin'
                    }

                };
            }


            // =================================================
            // ВЫРАВНИВАНИЕ
            // =================================================

            for (
                let rowNumber = 1;
                rowNumber <= sheet.rowCount;
                rowNumber++
            ) {

                sheet.getRow(
                    rowNumber
                ).alignment = {

                    vertical:
                        'middle',

                    wrapText:
                        true

                };
            }


            // =================================================
            // АВТОФИЛЬТР
            // =================================================

            sheet.autoFilter = {

                from:
                    'A1',

                to:
                    `G${dataLastRow}`

            };


            // =================================================
            // ЗАКРЕПЛЕНИЕ ШАПКИ
            // =================================================

            sheet.views = [

                {

                    state:
                        'frozen',

                    ySplit:
                        1

                }

            ];


            // =================================================
            // ПЕЧАТЬ
            //
            // Всё должно помещаться по ширине страницы.
            // =================================================

            sheet.pageSetup = {

                paperSize:
                    sheet.PAPERSIZE_A4,

                orientation:
                    'landscape',

                fitToPage:
                    true,

                fitToWidth:
                    1,

                fitToHeight:
                    0,

                horizontalDpi:
                    300,

                verticalDpi:
                    300,

                margins: {

                    left:
                        0.15,

                    right:
                        0.15,

                    top:
                        0.35,

                    bottom:
                        0.35,

                    header:
                        0.1,

                    footer:
                        0.1

                },

                horizontalCentered:
                    true

            };


            // =================================================
            // ПОВТОРЯТЬ ШАПКУ НА КАЖДОЙ СТРАНИЦЕ
            // =================================================

            sheet.pageSetup.printTitlesRow =
                '1:1';


            // =================================================
            // ОБЛАСТЬ ПЕЧАТИ
            // =================================================

            sheet.pageSetup.printArea =
                `A1:G${sheet.rowCount}`;


            // =================================================
            // КОЛОНТИТУЛ
            // =================================================

            sheet.headerFooter.oddHeader.center.text =
                `Отчёт по водителям: ${from} — ${to}`;


            sheet.headerFooter.oddHeader.center.size =
                8;


            sheet.headerFooter.oddFooter.center.text =
                'Страница &P из &N';


            sheet.headerFooter.oddFooter.center.size =
                8;


            // =================================================
            // XLSX
            // =================================================

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


            res.send(
                buffer
            );

        } catch (error) {

            console.error(
                'EXCEL ERROR:',
                error.response?.data ||
                error.message ||
                error
            );


            res
                .status(500)
                .json({

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

