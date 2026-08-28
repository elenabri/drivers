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
            'application/json;charset=utf-8'

    },

    timeout: 30000

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
// ПОЛУЧИТЬ ВСЕ ЗАПИСИ
// ============================================================

async function getAll(entity) {

    let offset = 0;

    const limit = 1000;

    const rows = [];


    while (true) {

        const response =
            await api.get(
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


        if (
            rows.length >= total
        ) {

            break;

        }


        if (
            !batch.length
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
        .replace(
            /\s+/g,
            ' '
        )
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
    // КОНТРАГЕНТЫ С ТЕГОМ "НАЕМНИКИ"
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


    const report = [];


    // ========================================================
    // ОБРАБОТКА СЧЕТОВ
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
        // Только выбранный период
        // ----------------------------------------------------

        if (
            invoiceDate < fromDate ||
            invoiceDate > toDate
        ) {

            continue;

        }


        // ====================================================
        // ПОЛУЧАЕМ ПОЛНЫЙ СЧЁТ
        //
        // Здесь дополнительно получаем state.
        // ====================================================

        let fullInvoice;


        try {

            fullInvoice =
                (
                    await api.get(
                        `/entity/invoicein/${invoice.id}?expand=agent,state`
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
        // КОНТРАГЕНТ
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
        // ТОЛЬКО КОНТРАГЕНТЫ С ТЕГОМ "НАЕМНИКИ"
        // ====================================================

        if (
            !mercenaryNames.has(
                normalizedCounterparty
            )
        ) {

            continue;

        }


        // ====================================================
        // СТАТУС СЧЁТА
        // ====================================================

        const invoiceStatus =
            fullInvoice.state?.name ||
            'Без статуса';


        // ====================================================
        // ДАТА
        // ====================================================

        const day =
            fullInvoice.moment.substring(
                0,
                10
            );


        // ====================================================
        // ОТГРУЗКИ
        // ====================================================

        const matchedDemands = [];


        let shipmentSum = 0;


        for (
            const demand
            of demands
        ) {

            if (
                !demand.moment
            ) {

                continue;

            }


            const demandDay =
                demand.moment.substring(
                    0,
                    10
                );


            // ------------------------------------------------
            // Отгрузка должна быть в тот же день
            // ------------------------------------------------

            if (
                demandDay !== day
            ) {

                continue;

            }


            // =================================================
            // ВОДИТЕЛЬ
            // =================================================

            const driverAttr =
                demand.attributes
                    ?.find(
                        a =>
                            a.name ===
                            'Водитель'
                    )
                    ?.value?.name ||
                '';


            if (
                !driverAttr
            ) {

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
                demand.agent?.name ||
                '';


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
                ) ===
                normalizedCounterparty;


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


            shipmentSum +=
                sum;


            matchedDemands.push({

                id:
                    demand.id,

                number:
                    demand.name,

                sum:
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
        // СУММА СЧЁТА
        // ====================================================

        const invoiceSum =
            fullInvoice.sum / 100;


        // ====================================================
        // 20%
        //
        // 20% от суммы отгрузок
        // ====================================================

        const twentyPercent =
            shipmentSum * 0.20;


        // ====================================================
        // РАЗНИЦА
        //
        // Сумма счета
        // -
        // Сумма отгрузок
        // -
        // 20%
        // ====================================================

        const difference =
            invoiceSum -
            shipmentSum -
            twentyPercent;


        // ====================================================
        // ДОБАВЛЯЕМ СТРОКУ
        // ====================================================

        report.push({

            invoiceId:
                fullInvoice.id,

            invoice:
                fullInvoice.name,

            // Статус нужен сайту для отображения
            // и фильтрации.
            //
            // В Excel мы его НЕ используем.
            status:
                invoiceStatus,

            date:
                day,

            counterparty:
                counterparty,

            driver:
                [
                    ...new Set(
                        matchedDemands
                            .map(
                                d =>
                                    d.driver
                            )
                            .filter(Boolean)
                    )
                ].join(', '),

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
            // Проверка дат
            // ------------------------------------------------

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
// API: EXCEL
//
// ВАЖНО:
//
// Excel получает УЖЕ ОТФИЛЬТРОВАННЫЕ строки.
//
// Статус счета здесь НЕ записываем.
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
            // Проверка
            // ------------------------------------------------

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
                            'Нет данных для Excel'

                    });

            }


            console.log(
                `Создание Excel: ${rows.length} строк`
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
            // КОЛОНКИ EXCEL
            //
            // СТАТУСА ЗДЕСЬ НЕТ.
            //
            // Также нет отдельного столбца
            // "Счет" и "Отгрузки".
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
                        13

                },

                {
                    header:
                        'Сумма счета',

                    key:
                        'invoiceSum',

                    width:
                        10

                },

                {
                    header:
                        'Сумма отгрузок',

                    key:
                        'shipmentSum',

                    width:
                        11

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
                        10

                }

            ];


            // =================================================
            // СТРОКИ
            // =================================================

            for (
                const row
                of rows
            ) {

                const invoiceSum =
                    Number(
                        row.invoiceSum ||
                        0
                    );


                const shipmentSum =
                    Number(
                        row.shipmentSum ||
                        0
                    );


                // ---------------------------------------------
                // 20%
                // ---------------------------------------------

                const twentyPercent =
                    shipmentSum *
                    0.20;


                // ---------------------------------------------
                // РАЗНИЦА
                // ---------------------------------------------

                const difference =
                    invoiceSum -
                    shipmentSum -
                    twentyPercent;


                sheet.addRow({

                    date:
                        row.date,

                    counterparty:
                        row.counterparty,

                    driver:
                        row.driver,

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
                sheet.getRow(
                    1
                );


            headerRow.font = {

                bold:
                    true

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
            // ИТОГО
            // =================================================

            totalRow.font = {

                bold:
                    true

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
            // ФОРМАТ ЧИСЕЛ
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
            // ВЫРАВНИВАНИЕ
            // =================================================

            for (
                let rowNumber = 1;
                rowNumber <= sheet.rowCount;
                rowNumber++
            ) {

                const row =
                    sheet.getRow(
                        rowNumber
                    );


                row.alignment = {

                    vertical:
                        'middle',

                    wrapText:
                        true

                };

            }


            // =================================================
            // АВТОФИЛЬТР
            //
            // ИТОГО в фильтр не входит.
            // =================================================

            sheet.autoFilter = {

                from:
                    'A1',

                to:
                    `G${dataLastRow}`

            };


            // =================================================
            // ЗАКРЕПИТЬ ЗАГОЛОВОК
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
            // Все колонки на одну ширину A4.
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
            // ПОВТОРЯТЬ ЗАГОЛОВОК
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
