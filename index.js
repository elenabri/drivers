```javascript
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

// В Vercel создай Environment Variable:
//
// MS_TOKEN = твой токен МойСклад
//
// Не храни токен непосредственно в коде.

const MS_TOKEN ='562d927aad09e55f49bed5ebd3751c7ccfd19a23';

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
// ПОЛУЧИТЬ ВСЕ ЗАПИСИ
//
// Запросы выполняются последовательно.
// Это важно, чтобы не создавать лишнюю нагрузку
// на API МойСклад.
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
        .replace(
            /\s+/g,
            ' '
        )
        .toLowerCase();

}


// ============================================================
// ПОЛУЧЕНИЕ АТРИБУТА
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
// ВАЖНО:
// Оставляем именно ту логику, которая работала раньше.
//
// demand.attributes -> "Водитель"
// ============================================================

function getDriverFromDemand(
    demand
) {

    const attribute =
        demand.attributes
            ?.find(
                a =>
                    a.name ===
                    'Водитель'
            );


    return (
        attribute?.value?.name ||
        ''
    );

}


// ============================================================
// СУММА МОЙСКЛАД
//
// В МойСклад сумма хранится в копейках.
// ============================================================

function money(value) {

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
// СОПОСТАВЛЕНИЕ ОТГРУЗКИ СО СЧЁТОМ
//
// ЭТУ ЛОГИКУ НЕ МЕНЯЕМ.
//
// Отгрузка подходит, если:
//
// 1. водитель входит в название контрагента счета
//
// ИЛИ
//
// 2. контрагент отгрузки полностью совпадает
//    с контрагентом счета
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
        demand.agent?.name ||
        '';


    const normalizedDemandCounterparty =
        normalizeName(
            demandCounterparty
        );


    // --------------------------------------------------------
    // Водитель входит в название контрагента
    // --------------------------------------------------------

    const driverMatches =
        normalizedCounterparty.includes(
            normalizedDriver
        );


    // --------------------------------------------------------
    // Контрагент отгрузки = контрагент счета
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
        `ОТЧЁТ ПО ВОДИТЕЛЯМ: ${from} -> ${to}`
    );


    console.log(
        '=========================================='
    );


    // ========================================================
    // ЗАГРУЖАЕМ ДАННЫЕ ПОСЛЕДОВАТЕЛЬНО
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


        // ----------------------------------------------------
        // ДАТА СЧЁТА
        // ----------------------------------------------------

        const invoiceDateString =
            invoice.moment.substring(
                0,
                10
            );


        // ----------------------------------------------------
        // ПЕРИОД
        // ----------------------------------------------------

        if (
            invoiceDateString < from ||
            invoiceDateString > to
        ) {

            continue;

        }


        // ====================================================
        // ПОЛНЫЙ СЧЁТ
        //
        // Получаем:
        // agent
        // state
        // ====================================================

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
        // ДАТА СЧЁТА
        // ====================================================

        const day =
            fullInvoice.moment.substring(
                0,
                10
            );


        // ====================================================
        // СТАТУС СЧЁТА
        // ====================================================

        const status =
            getInvoiceStatus(
                fullInvoice
            );


        // ====================================================
        // ПОДХОДЯЩИЕ ОТГРУЗКИ
        // ====================================================

        const matchedDemands = [];


        let shipmentSum =
            0;


        // ====================================================
        // ПРОХОДИМ ПО ВСЕМ ОТГРУЗКАМ
        // ====================================================

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
            // ДАТА ОТГРУЗКИ
            // ------------------------------------------------

            const demandDay =
                demand.moment.substring(
                    0,
                    10
                );


            // ------------------------------------------------
            // ТОЛЬКО ТОТ ЖЕ ДЕНЬ
            // ------------------------------------------------

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
            // НЕ МЕНЯЕМ.
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
            // СУММА ОТГРУЗКИ
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
            // Сохраняем конкретную отгрузку
            // ------------------------------------------------

            matchedDemands.push({

                id:
                    demand.id,

                number:
                    demand.name ||
                    '',

                sum:
                    sum,

                driver:
                    driverAttr,

                counterparty:
                    demand.agent?.name ||
                    ''

            });

        }


        // ====================================================
        // ЕСЛИ ОТГРУЗОК НЕТ — НЕ ПОКАЗЫВАЕМ СЧЁТ
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
            shipmentSum *
            0.20;


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
        // ВОДИТЕЛИ
        //
        // ТОЛЬКО ИЗ matchedDemands
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
        // ДОБАВЛЯЕМ СТРОКУ
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
// API: ПОЛУЧИТЬ ОТЧЁТ
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


            // ------------------------------------------------
            // Проверяем даты
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
                            'Дата начала не может быть позже даты окончания'

                    });

            }


            // ------------------------------------------------
            // Формируем отчёт
            // ------------------------------------------------

            const report =
                await buildDriverReport(
                    from,
                    to
                );


            // ------------------------------------------------
            // Возвращаем JSON
            // ------------------------------------------------

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
// Здесь МойСклад НЕ вызывается.
//
// В Excel приходит уже отфильтрованный список.
//
// Статус счета сюда НЕ попадает.
// Счёт сюда НЕ попадает.
// Детализация отгрузок сюда НЕ попадает.
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


            // ------------------------------------------------
            // Проверка периода
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


            // ------------------------------------------------
            // Проверка строк
            // ------------------------------------------------

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
            // КОЛОНКИ EXCEL
            //
            // 1 Дата
            // 2 Контрагент
            // 3 Водитель
            // 4 Сумма счета
            // 5 Сумма отгрузок
            // 6 20%
            // 7 Разница
            //
            // НЕ ДОБАВЛЯЕМ:
            // Счёт
            // Статус
            // Отгрузки
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
            // СТРОКИ
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
                    shipmentSum *
                    0.20;


                // ------------------------------------------------
                // Разница
                // ------------------------------------------------

                const difference =
                    invoiceSum -
                    shipmentSum -
                    twentyPercent;


                sheet.addRow({

                    date:
                        row.date ||
                        '',

                    counterparty:
                        row.counterparty ||
                        '',

                    driver:
                        row.driver ||
                        '',

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

            const headerRow =
                sheet.getRow(
                    1
                );


            headerRow.font = {

                bold:
                    true,

                size:
                    10

            };


            headerRow.alignment = {

                horizontal:
                    'center',

                vertical:
                    'middle',

                wrapText:
                    true

            };


            headerRow.height =
                25;


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

                totalRow
                    .getCell(col)
                    .border = {

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

                sheet
                    .getRow(rowNumber)
                    .alignment = {

                        vertical:
                            'middle',

                        wrapText:
                            true

                    };

            }


            // =================================================
            // ЧИСЛА СПРАВА
            // =================================================

            for (
                let rowNumber = 2;
                rowNumber <= sheet.rowCount;
                rowNumber++
            ) {

                for (
                    const column
                    of ['D', 'E', 'F', 'G']
                ) {

                    sheet
                        .getCell(
                            `${column}${rowNumber}`
                        )
                        .alignment = {

                            vertical:
                                'middle',

                            horizontal:
                                'right'

                        };

                }

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
            // ЗАКРЕПЛЯЕМ ШАПКУ
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
            // A4
            // Альбомная
            // Все колонки по ширине одной страницы
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
                        0.30,

                    bottom:
                        0.30,

                    header:
                        0.10,

                    footer:
                        0.10

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
```
