
const express = require('express');
const axios = require('axios');
const ExcelJS = require('exceljs');

const app = express();

app.use(express.json({
    limit: '10mb'
});


// ============================================================
// MOYSKLAD
// ============================================================

const MS_TOKEN = '562d927aad09e55f49bed5ebd3751c7ccfd19a23';

if (!MS_TOKEN) {
    console.warn('WARNING: MS_TOKEN не задан');
}

const api = axios.create({

    baseURL:
        'https://api.moysklad.ru/api/remap/1.2',

    headers: {

        Authorization:
            `Bearer ${MS_TOKEN}`,

        Accept:
            'application/json;charset=utf-8',

        'Accept-Encoding':
            'gzip'

    },

    timeout: 60000

});


// ============================================================
// ГЛАВНАЯ
// ============================================================

app.get('/', (req, res) => {

    res.sendFile(
        __dirname + '/index.html'
    );

});


// ============================================================
// ПОЛУЧЕНИЕ ВСЕХ ЗАПИСЕЙ
//
// ВАЖНО:
// Запросы идут последовательно.
// Не используем Promise.all для МойСклад,
// чтобы не ловить ошибку 1073.
// ============================================================

async function getAll(entity) {

    let offset = 0;

    const limit = 1000;

    const rows = [];

    while (true) {

        const response =
            await api.get(
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
                response.data?.meta?.size || 0
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
// НОРМАЛИЗАЦИЯ
// ============================================================

function normalizeName(name) {

    return String(
        name || ''
    )
        .trim()
        .replace(/\s+/g, ' ')
        .toLowerCase();

}


// ============================================================
// АТРИБУТ
// ============================================================

function getAttributeValue(
    entity,
    attributeName
) {

    const attribute =
        entity?.attributes?.find(
            a =>
                normalizeName(a.name) ===
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
// ВОДИТЕЛЬ ИЗ ОТГРУЗКИ
//
// Это принципиально важно.
//
// Берём именно:
// demand.attributes -> "Водитель"
// ============================================================

function getDriverFromDemand(
    demand
) {

    const attribute =
        demand.attributes?.find(
            a =>
                a.name === 'Водитель'
        );

    return (
        attribute?.value?.name ||
        ''
    );

}


// ============================================================
// СУММА MOYSKLAD
// ============================================================

function money(
    value
) {

    return Number(
        value || 0
    ) / 100;

}


// ============================================================
// СТАТУС СЧЁТА
// ============================================================

function getInvoiceStatus(
    invoice
) {

    return (
        invoice?.state?.name ||
        'Без статуса'
    );

}


// ============================================================
// ПРОВЕРКА:
// ЯВЛЯЕТСЯ ЛИ КОНТРАГЕНТ НАЕМНИКОМ
// ============================================================

function isMercenary(
    counterparty,
    mercenaryNames
) {

    return mercenaryNames.has(
        normalizeName(
            counterparty
        )
    );

}


// ============================================================
// ОСНОВНОЕ СОПОСТАВЛЕНИЕ
//
// ЭТО НЕ МЕНЯЕМ.
//
// Старая логика:
// 1. Отгрузка в тот же день.
// 2. Есть водитель.
// 3. Водитель содержится в названии контрагента счета.
//    ИЛИ
// 4. Контрагент отгрузки полностью совпадает
//    с контрагентом счета.
// ============================================================

function demandMatchesInvoice(
    demand,
    invoiceCounterparty
) {

    const driverAttr =
        getDriverFromDemand(
            demand
        );

    if (!driverAttr) {
        return false;
    }

    const normalizedDriver =
        normalizeName(
            driverAttr
        );

    const normalizedCounterparty =
        normalizeName(
            invoiceCounterparty
        );

    const demandCounterparty =
        demand.agent?.name || '';

    const normalizedDemandCounterparty =
        normalizeName(
            demandCounterparty
        );


    // --------------------------------------------------------
    // ГЛАВНАЯ ПРОВЕРКА ПО ВОДИТЕЛЮ
    // --------------------------------------------------------

    const driverMatches =
        normalizedCounterparty.includes(
            normalizedDriver
        );


    // --------------------------------------------------------
    // ДОПОЛНИТЕЛЬНОЕ СОВПАДЕНИЕ ПО КОНТРАГЕНТУ
    // --------------------------------------------------------

    const counterpartyMatches =
        normalizedDemandCounterparty ===
        normalizedCounterparty;


    return (
        driverMatches ||
        counterpartyMatches
    );

}


// ============================================================
// ПОСТРОЕНИЕ ОТЧЁТА
// ============================================================

async function buildDriverReport(
    from,
    to
) {

    console.log(
        '=========================================='
    );

    console.log(
        `DRIVER REPORT: ${from} -> ${to}`
    );

    console.log(
        '=========================================='
    );


    // ========================================================
    // НЕ ДЕЛАЕМ Promise.all
    //
    // Чтобы не превышать лимит МойСклад.
    // ========================================================

    const invoices =
        await getAll(
            'invoicein'
        );

    const demands =
        await getAll(
            'demand'
        );

    const counterparties =
        await getAll(
            'counterparty'
        );


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
    // КОНТРАГЕНТЫ "НАЕМНИКИ"
    //
    // Оставляем именно старую проверку tags.
    // ========================================================

    const mercenaryNames =
        new Set();


    for (
        const cp
        of counterparties
    ) {

        if (
            cp.tags &&
            cp.tags.includes(
                'наемники'
            )
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
    // ОТЧЁТ
    // ========================================================

    const report = [];


    // ========================================================
    // ОБРАБАТЫВАЕМ СЧЕТА
    // ========================================================

    for (
        const invoice
        of invoices
    ) {

        if (
            !invoice.moment
        ) {
            continue;
        }


        const invoiceDate =
            new Date(
                invoice.moment
            );


        // ----------------------------------------------------
        // Проверяем период
        // ----------------------------------------------------

        const invoiceDateString =
            invoice.moment.substring(
                0,
                10
            );


        if (
            invoiceDateString < from ||
            invoiceDateString > to
        ) {
            continue;
        }


        // ----------------------------------------------------
        // Получаем полный счёт
        //
        // Теперь одновременно получаем:
        // agent
        // state
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
                'Ошибка получения счета:',
                invoice.id,

                error.response?.data ||
                error.message
            );

            continue;
        }


        // ====================================================
        // КОНТРАГЕНТ СЧЁТА
        // ====================================================

        if (
            !fullInvoice.agent?.name
        ) {
            continue;
        }


        const counterparty =
            fullInvoice.agent.name;


        const normalizedCounterparty =
            normalizeName(
                counterparty
            );


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


        // ====================================================
        // ДАТА
        // ====================================================

        const day =
            fullInvoice.moment.substring(
                0,
                10
            );


        // ====================================================
        // СТАТУС
        // ====================================================

        const status =
            getInvoiceStatus(
                fullInvoice
            );


        // ====================================================
        // СОБИРАЕМ ТОЛЬКО ПОДХОДЯЩИЕ ОТГРУЗКИ
        // ====================================================

        const matchedDemands = [];


        let shipmentSum =
            0;


        for (
            const demand
            of demands
        ) {

            if (
                !demand.moment
            ) {
                continue;
            }


            // ------------------------------------------------
            // ТОЛЬКО ТОТ ЖЕ ДЕНЬ
            // ------------------------------------------------

            const demandDay =
                demand.moment.substring(
                    0,
                    10
                );


            if (
                demandDay !== day
            ) {
                continue;
            }


            // ------------------------------------------------
            // ВОДИТЕЛЬ
            // ------------------------------------------------

            const driverAttr =
                getDriverFromDemand(
                    demand
                );


            if (
                !driverAttr
            ) {
                continue;
            }


            // ------------------------------------------------
            // СОПОСТАВЛЕНИЕ
            //
            // ИМЕННО СТАРАЯ ЛОГИКА.
            // ------------------------------------------------

            if (
                !demandMatchesInvoice(
                    demand,
                    counterparty
                )
            ) {
                continue;
            }


            // ------------------------------------------------
            // СУММА
            // ------------------------------------------------

            if (
                (demand.sum || 0) <= 0
            ) {
                continue;
            }


            const sum =
                money(
                    demand.sum
                );


            shipmentSum +=
                sum;


            // ------------------------------------------------
            // Сохраняем КОНКРЕТНУЮ отгрузку
            // ------------------------------------------------

            matchedDemands.push({

                id:
                    demand.id,

                number:
                    demand.name || '',

                sum:
                    sum,

                driver:
                    driverAttr,

                counterparty:
                    demand.agent?.name || ''

            });

        }


        // ====================================================
        // ЕСЛИ НЕТ ОТГРУЗОК — СЧЁТ НЕ ПОКАЗЫВАЕМ
        // ====================================================

        if (
            matchedDemands.length === 0
        ) {
            continue;
        }


        // ====================================================
        // СУММА СЧЁТА
        // ====================================================

        const invoiceSum =
            money(
                fullInvoice.sum
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
        //
        // = 4000
        // ====================================================

        const difference =
            invoiceSum -
            shipmentSum -
            twentyPercent;


        // ====================================================
        // ТОЛЬКО РЕАЛЬНЫЕ ВОДИТЕЛИ ИЗ matchedDemands
        // ====================================================

        const drivers =
            [
                ...new Set(
                    matchedDemands
                        .map(
                            d =>
                                d.driver
                        )
                        .filter(Boolean)
                )
            ];


        // ====================================================
        // СТРОКА ОТЧЁТА
        // ====================================================

        report.push({

            invoiceId:
                fullInvoice.id,

            invoice:
                fullInvoice.name ||
                '',

            status:
                status,

            date:
                day,

            counterparty:
                counterparty,

            driver:
                drivers.join(
                    ', '
                ),

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
// API: ОТЧЁТ
// ============================================================

app.get(
    '/api/driver-report',
    async (
        req,
        res
    ) => {

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


            return res.json({

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


            return res
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
// API: EXCEL
//
// ВАЖНО:
//
// МойСклад здесь вообще НЕ вызывается.
//
// Получаем уже отфильтрованный filteredReport
// из браузера.
//
// Поэтому при скачивании не должно быть новых
// запросов к МойСклад.
// ============================================================

app.post(
    '/api/driver-report.xlsx',
    async (
        req,
        res
    ) => {

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
            // SHEET
            // =================================================

            const sheet =
                workbook.addWorksheet(
                    'Отчет'
                );


            // =================================================
            // EXCEL-КОЛОНКИ
            //
            // НЕТ:
            // Счёта
            // Статуса
            // Отгрузок
            //
            // Есть только:
            // Дата
            // Контрагент
            // Водитель
            // Сумма счета
            // Сумма отгрузок
            // 20%
            // Разница
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


                // ------------------------------------------------
                // 20%
                // ------------------------------------------------

                const twentyPercent =
                    shipmentSum * 0.20;


                // ------------------------------------------------
                // РАЗНИЦА
                // ------------------------------------------------

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
            // ШАПКА
            // =================================================

            const header =
                sheet.getRow(
                    1
                );


            header.font = {

                bold:
                    true,

                size:
                    10

            };


            header.alignment = {

                horizontal:
                    'center',

                vertical:
                    'middle',

                wrapText:
                    true

            };


            header.height =
                25;


            // =================================================
            // ФОРМАТ ЧИСЕЛ
            // =================================================

            for (
                let i = 2;
                i <= sheet.rowCount;
                i++
            ) {

                sheet.getCell(
                    `D${i}`
                ).numFmt =
                    '#,##0.00';


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

            }


            // =================================================
            // ИТОГО
            // =================================================

            totalRow.font = {

                bold:
                    true

            };


            totalRow.height =
                23;


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
                let i = 1;
                i <= sheet.rowCount;
                i++
            ) {

                sheet.getRow(
                    i
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
            // ЗАКРЕПИТЬ ШАПКУ
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
            // Всё по ширине страницы.
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
                        0.3,

                    bottom:
                        0.3,

                    header:
                        0.1,

                    footer:
                        0.1

                },

                horizontalCentered:
                    true

            };


            // =================================================
            // ПОВТОР ШАПКИ
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
            // ФОРМИРУЕМ XLSX
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


            return res.send(
                buffer
            );

        } catch (error) {

            console.error(
                'EXCEL ERROR:',
                error.response?.data ||
                error.message ||
                error
            );


            return res
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

