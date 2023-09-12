// 導入所需的庫
const request = require('request');
const cheerio = require('cheerio');
const util = require('util');
const ExcelJS = require('exceljs');  // 新增：引入 exceljs 庫

// 將 request 函數轉換為基於 Promise 的函數
const requestPromise = util.promisify(request);

// 帶有重試邏輯的 URL 抓取函數
async function fetchWithRetry(url, delayInSeconds = 1, retryCount = 0) {
  console.log(`正在嘗試抓取 ${url}, 重試次數：${retryCount}`);

  try {
    const { error, statusCode, body } = await requestPromise(url);
    if (!error && statusCode === 200) {
      return body;
    } else {
      retryCount++;
      console.log(`抓取 ${url} 失敗。狀態碼：${statusCode}。${retryCount} 次重試，延遲 ${delayInSeconds} 秒...`);
      await new Promise(resolve => setTimeout(resolve, delayInSeconds * 1000));
      return await fetchWithRetry(url, delayInSeconds + 1, retryCount);
    }
  } catch (err) {
    retryCount++;
    console.error(`抓取 ${url} 失敗：${err.message}。${retryCount} 次重試，延遲 ${delayInSeconds} 秒...`);
    await new Promise(resolve => setTimeout(resolve, delayInSeconds * 1000));
    return await fetchWithRetry(url, delayInSeconds + 1, retryCount);
  }
}

// 主函數：抓取卡片圖片並保存數據
async function fetchCardImages() {
  console.log("開始抓取URL...");
  const urls = [];
  const cardDataArray = [['imgSrc', 'cardNameJapanese', 'set', 'rarity', 'cardNumber']];

  // 用於追蹤已訪問過的卡號的 Set
  const visitedCardNos = new Set();

  // 填充 URLs
  urls.push('https://ws-tcg.com/cardlist/search');
  for (let i = 2; i <= 1786; i++) {
    urls.push(`https://ws-tcg.com/cardlist/search?page=${i}`);
  }

  // 遍歷每個 URL
  for (let url of urls) {
    console.log(`正在抓取 ${url}...`);
    const body = await fetchWithRetry(url);
    const $ = cheerio.load(body);
    
    // 在頁面上抓取卡片
    const cards = $('a[href^="/cardlist/?cardno="]').toArray();
    
    for (let el of cards) {
      const cardNo = $(el).attr('href').split('=')[1].split('&')[0];
      if (visitedCardNos.has(cardNo)) {
        continue;
      }
      visitedCardNos.add(cardNo);

      const cardUrl = `https://ws-tcg.com/cardlist/?cardno=${cardNo}`;
      const cardBody = await fetchWithRetry(cardUrl);
      const $$ = cheerio.load(cardBody);
      
      const imgShortSrc = $$('tbody tr:first-child td:first-child img').attr('src');
      const imgSrc = `https://ws-tcg.com${imgShortSrc}`;
      const cardNameJapanese = $$('tbody tr:first-child td:nth-child(3)').contents().filter(function() {
        return this.nodeType === 3;
      }).text().trim();
      const set = $$('tbody tr:nth-child(3) td').html();
      const rarity = $$('tbody tr:nth-child(5) td:nth-child(4)').text();
      const cardNumber = $$('tbody tr:nth-child(2) td').html();
      
      cardDataArray.push([imgSrc, cardNameJapanese, set, rarity, cardNumber]);
    }
  }

  // 新增：創建 Excel 工作簿和工作表
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Card Data');
  
  // 新增：設置表頭
  worksheet.columns = [
    { header: 'Produc tId', key: 'ProductId'},
    { header: 'Set', key: 'set' },
    { header: 'Edition', key: 'edition' },
    { header: 'Series', key: 'series' },
    { header: 'Rarity', key: 'rarity' },
    { header: 'Material', key: 'material' },
    { header: 'ReleaseYear', key: 'release year' },
    { header: 'Language', key: 'language' },
    { header: 'Card Name English', key: 'cardName english' },
    { header: 'Card Name Chinese', key: 'cardName chinese' },
    { header: 'Card Name (Japanese)', key: 'cardNameJapanese' },
    { header: 'ReleaseYear', key: 'release year' },
    { header: 'Card Number', key: 'cardNumber' },
    { header: 'Img Src', key: 'imgSrc' },
    { header: 'Value', key: 'value' },
    { header: 'Reference', key: 'reference' },
    { header: 'Remark', key: 'remark' },
    { header: 'Remark1', key: 'remark1' },
    { header: 'Remark2', key: 'remark2' },
    { header: 'Remark3', key: 'remark3' },
    { header: 'Remark4', key: 'remark4' },
    { header: 'Remark5', key: 'remark5' },
    { header: 'Remark6', key: 'remark6' },
    { header: 'Remark7', key: 'remark7' },
    { header: 'Remark8', key: 'remark8' },
    { header: 'Remark9', key: 'remark9' },
    { header: 'Remark10', key: 'remark10' },
    { header: 'Enable', key: 'enable' },
    { header: 'P_Language', key: 'pLanguage' },
    { header: 'Id', key: 'id' },
  ];
  
  // 新增：添加行
// 新增：添加行
for (const data of cardDataArray.slice(1)) {
    worksheet.addRow({
      imgSrc: data[0],
      cardNameJapanese: data[1],
      set: data[2],
      rarity: data[3],
      cardNumber: data[4],
      // 其他都為空值
      edition: null,
      series: null,
      material: null,
      releaseYear: null,
      language: null,
      cardNameEnglish: null,
      cardNameChinese: null,
      value: null,
      reference: null,
      remark: null,
      remark1: null,
      remark2: null,
      remark3: null,
      remark4: null,
      remark5: null,
      remark6: null,
      remark7: null,
      remark8: null,
      remark9: null,
      remark10: null,
      enable: null,
      pLanguage: null,
      id: null
    });
  }
  

  // 新增：保存為 Excel 文件
  await workbook.xlsx.writeFile('cardData.xlsx');

  console.log("完成卡片圖片收集，數據已保存為 Excel 文件。");
}

// 執行主函數
fetchCardImages();