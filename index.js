const TelegramApi = require('node-telegram-bot-api');
const xlsx = require('xlsx');
require('dotenv').config();
const path = require('path');
const express = require('express');
const { gameOptions, againOptions } = require('./options');

const token = process.env.TELEGRAM_BOT_TOKEN;
const bot = new TelegramApi(token, { polling: true });

const app = express();
const port = process.env.PORT || 3000;
const chats = {};
const chatState = {};

const fileNames = ['prokat.xlsx', 'umka.xlsx', 'svatik.xlsx'];

// Функція для завантаження та обробки Excel-файлу
const loadExcelFile = (fileName) => {
    const filePath = path.join(__dirname, fileName);
    let workbook;
    try {
        workbook = xlsx.readFile(filePath);
        console.log(`Файл ${fileName} завантажено успішно`);
    } catch (err) {
        console.error(`Помилка при завантаженні файлу ${fileName}:`, err);
        return [];
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    if (data.length < 4) {
        console.error(`Файл ${fileName} має недостатньо рядків для обробки.`);
        return [];
    }

    const headers = data[2];
    const rows = data.slice(3).map(row => {
        let rowObject = {};
        headers.forEach((header, index) => {
            rowObject[header] = row[index];
        });
        return { ...rowObject, sourceFile: fileName };
    });

    console.log(`Завантажено ${rows.length} рядків з файлу ${fileName}`);
    return rows;
};

const loadAllExcelFiles = () => {
    return fileNames.flatMap(fileName => loadExcelFile(fileName));
};

let rows = loadAllExcelFiles();

const startGame = async (chatId) => {
    await bot.sendMessage(chatId, 'Зараз я загадаю число від 0 до 9 а ти відгадуй');
    const randomNumber = Math.floor(Math.random() * 10);
    chats[chatId] = randomNumber;
    await bot.sendMessage(chatId, 'Відгадуй', gameOptions);
};

const sendLongMessage = async (chatId, message) => {
    const chunkSize = 4096;
    for (let i = 0; i < message.length; i += chunkSize) {
        await bot.sendMessage(chatId, message.substring(i, i + chunkSize));
    }
};

const start = () => {
    bot.setMyCommands([
        { command: '/start', description: 'Початкове привітання' },
        { command: '/info', description: 'Получити інформацію про користувача' },
        { command: '/game', description: 'Гра вгадай цифру' },
        { command: '/kod', description: 'Пошук штрихкоду' },
        { command: '/article', description: 'Пошук за артикулом' },
        { command: '/name', description: 'Пошук за назвою' },  
    ]);

    bot.on('message', async (msg) => {
        const text = msg.text;
        const chatId = msg.chat.id;

        if (text === '/start') {
            await bot.sendMessage(chatId, `Щоб почати користуватися ботом треба тикнути цю кнопочку`);
            await bot.sendPhoto(chatId, path.join(__dirname, 'start_1.png'));
            await bot.sendMessage(chatId, `ДУМАЙТЕ ЩО ПИШИТЕ В ПОШУК ПО НАЗВІ`);
            await bot.sendMessage(chatId, `Не пишіть просто "штани" чи щось бо буде дуже спамити повідомленями таке краще вписуйте модель!!!!`);
            await bot.sendPhoto(chatId, path.join(__dirname, 'nameSearch.png'));
            await bot.sendMessage(chatId, `Назви складів umka - Умка, svatik - святік, prokat - прокат`);
            await bot.sendMessage(chatId, `Вітаю вас в базі даних UMKA.SHOP`);
            return bot.sendSticker(chatId, 'https://sl.combot.org/tommood_by_moe_sticker_bot/webp/33xf09f9299.webp');
        }

        if (text === '/info') {
            return bot.sendMessage(chatId, `Тебе звати ${msg.from.first_name} ${msg.from.last_name}`);
        }

        if (text === '/game') {
            return startGame(chatId);
        }

        if (text === '/kod' || text === '/article' || text === '/name') {
            await bot.sendMessage(chatId, text === '/kod' ? 'Введіть штрих-код:' : text === '/article' ? 'Введіть артикул:' : 'Введіть назву товару:');
            chatState[chatId] = text === '/kod' ? 'awaiting_barcode' : text === '/article' ? 'awaiting_article' : 'awaiting_name';
            return;
        }

        if (chatState[chatId] === 'awaiting_barcode') {
            const barcode = msg.text;
            const result = rows.find(row => String(row['Штрих-код']) === String(barcode));

            if (result) {
                const article = result['Артикул'];
                const relatedItems = rows.filter(row => row['Артикул'] === article);

                let response = `Знайдено товар у файлі ${result.sourceFile}:\nНазва товару: ${result['Назва товару']}\nШтрих-код: ${result['Штрих-код']}\nЦіна роздрібна: ${result['Ціна роздрібна']}\nКількість: ${result['Кількість']}\nАртикул: ${result['Артикул']}\n\n`;
                if (relatedItems.length > 1) {
                    response += 'Cхожі товари:\n';
                    relatedItems.forEach(item => {
                        if (item['Штрих-код'] !== barcode) {
                            response += `Назва товару: ${item['Назва товару']}\nШтрих-код: ${item['Штрих-код']}\nЦіна роздрібна: ${item['Ціна роздрібна']}\nКількість: ${item['Кількість']}\nАртикул: ${item['Артикул']}\nСклад: ${item.sourceFile}\n\n`;
                        }
                    });
                }
                await sendLongMessage(chatId, response);
            } else {
                await bot.sendMessage(chatId, 'Штрих-код не знайдено');
            }

            chatState[chatId] = null;
            return;
        }

        if (chatState[chatId] === 'awaiting_article') {
            const article = msg.text;
            const results = rows.filter(row => String(row['Артикул']) === String(article));

            if (results.length > 0) {
                let response = 'Знайдено товари:\n';
                results.forEach(result => {
                    response += `Назва товару: ${result['Назва товару']}\nШтрих-код: ${result['Штрих-код']}\nЦіна роздрібна: ${result['Ціна роздрібна']}\nКількість: ${result['Кількість']}\nАртикул: ${result['Артикул']}\nСклад: ${result.sourceFile}\n\n`;
                });
                await sendLongMessage(chatId, response);
            } else {
                await bot.sendMessage(chatId, 'Артикул не знайдено');
            }

            chatState[chatId] = null;
            return;
        }

        if (chatState[chatId] === 'awaiting_name') {
            const name = msg.text.toLowerCase();
            const results = rows.filter(row => String(row['Назва товару']).toLowerCase().includes(name));

            if (results.length > 0) {
                let response = 'Знайдено товари:\n';
                results.forEach(result => {
                    response += `Назва товару: ${result['Назва товару']}\nШтрих-код: ${result['Штрих-код']}\nЦіна роздрібна: ${result['Ціна роздрібна']}\nКількість: ${result['Кількість']}\nАртикул: ${result['Артикул']}\nСклад: ${result.sourceFile}\n\n`;
                });
                await sendLongMessage(chatId, response);
            } else {
                await bot.sendMessage(chatId, 'Товар не знайдено');
            }

            chatState[chatId] = null;
            return;
        }

        return bot.sendMessage(chatId, 'Я тебе не розумію спробуй ще раз!');
    });

    bot.on('callback_query', async (msg) => {
        const data = msg.data;
        const chatId = msg.message.chat.id;

        if (data === '/again') {
            return startGame(chatId);
        }

        if (Number(data) === chats[chatId]) {
            await bot.sendMessage(chatId, `Вітаю ти відгадав цифру ${chats[chatId]}`, againOptions);
            return bot.sendSticker(chatId, 'https://sl.combot.org/peachngomavnt/webp/55xf09f8e89.webp');
        } else {
            return await bot.sendMessage(chatId, `Нажаль ти не вгадав цифру, бот загадав цифру: ${chats[chatId]}`, againOptions);
        }
    });
};

start();

app.get('/', (req, res) => {
    console.log('Received request');
    res.send('Bot is running');
});

app.listen(port, () => {
    console.log(`Server is listening on port ${port}`);
});
