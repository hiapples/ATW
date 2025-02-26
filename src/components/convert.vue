<script setup>
import { ref, nextTick } from 'vue';
import ExcelJS from 'exceljs';

const inputText = ref(''); // 輸入框内容
const outputText = ref(''); // 輸出结果
const waybillCount = ref(''); //運單數量
const notificationMessage = ref(''); // 通知消息
const notificationType = ref(''); // 通知類型
const showNotification = ref(false); // 控制通知顯示
let notificationTimeout; // 計時器参考

// 顯示通知
const showNotificationWithDelay = (message, type) => {
  if (notificationTimeout) clearTimeout(notificationTimeout);
  showNotification.value = false;

  setTimeout(() => {
    notificationMessage.value = message.replace(/\n/g, '<br>'); // 替換換行符為 <br> 標籤
    notificationType.value = type;
    showNotification.value = true;
    notificationTimeout = setTimeout(() => {
      showNotification.value = false;
    }, 5000); // 顯示通知 5 秒後關閉
  }, 300);
};

// 固定列數檢查
const validateLine = (line) => {
  const parts = line.split(/\t+/); // 按 Tab 分隔

  if (parts.length < 11) {
    outputText.value = '';
    showModal.value = false;
    return 'Error: Too few columns (less than 11)';
  }
  //客貨訂單號
  const orderNumber = parts[0]; // 銷貨單號
  const orderNumberRegex = /^[a-zA-Z0-9#]*$/; // 正規表達式，要求只包含英文和数字和#

  if (!orderNumberRegex.test(orderNumber)) {
    outputText.value = '';
    showModal.value = false;
    return 'Error: Order number must contain only letters, digits, or #';
  }
  //電話
  let phone = parts[4].trim().replace(/\s+/g, ''); // 去除前後空格並刪除所有空格
  const phoneRegex = /^[\d#]+$/; // 只允許數字和 #

  if (!phoneRegex.test(phone)) {
    outputText.value = '';
    showModal.value = false;
    return 'Error: Phone number must contain only digits or #';
  }
  // 地址
  const address = normalizeAddress(parts[5]); 
  const validCities = [
    '新竹市', '台南市', '台北市', '高雄市', '桃園市', 
    '台中市', '彰化縣', '嘉義市', '屏東縣', '雲林縣',
    '基隆市', '宜蘭縣', '新北市', '南投縣', '嘉義縣',
    '花蓮縣', '台東縣', '連江縣', '金門縣', '澎湖縣'
  ];
  // 將有效城市轉換為正規表達式
  const cityPattern = validCities.join('|'); // 將城市名用 | 連接成正規表達式
  const regex = new RegExp(`(${cityPattern})`, 'g'); // 創建全局正規表達式

  // 找到地址中的所有縣市
  const foundCities = address.match(regex);

  // 檢查是否有多個城市
  if (foundCities && foundCities.length > 1) {
    outputText.value = '';
    showModal.value = false;
    return 'Error: Address must contain exactly one city/county';
  }


  
  //商品數量
  const quantity = parts[10]; 
  const quantityRegex = /^[1-9]\d*$/; // 正規表達式，要求只包含数字

  if (!quantityRegex.test(quantity)) {
    outputText.value = '';
    showModal.value = false;
    return 'Error: Quantity must be a number';
  }

  if (parts.length > 11) {
    outputText.value = '';
    showModal.value = false;
    return 'Error: Too many columns (greater than 11)';
  }


  return null; // 如果没有錯誤
};

//輸入處理
const handleInput = () => {
  // 使用正則表達式將兩個 Tab 中間加上空格
  inputText.value = inputText.value.replace(/\t\t/g, '\t \t');
  // 按換行符拆分輸入内容為行
  let lines = inputText.value.split('\n');

  // 初始化结果數组
  let result = [];
  let tempLine = '';

  lines.forEach((line) => {
    // 去掉前後空格
    line = line.trim();

    // 檢查行是否以數字結尾（即是完整紀錄）
    if (/.*\d$/.test(line)) {
      // 如果當前有臨時行，合併到結果中
      if (tempLine) {
        tempLine += ' ' + line; // 合併紀錄
        result.push(tempLine.trim());
        tempLine = ''; // 清空臨時行
      } else {
        // 如果没有臨時行，直接加入結果
        result.push(line);
      }
    } else {
      // 行内數據需要合併到臨時行
      tempLine += ' ' + line;
    }
  });

  // 如果最後有剩餘的臨時行，加入结果
  if (tempLine) {
    result.push(tempLine.trim());
  }

  // 將結果數组合併為字符串，保留紀錄間的換行符
  inputText.value = result.join('\n');
};

//地址標準化
function normalizeAddress(address) {
  return address
    .replace(/^台灣\s*\d{3}\s*/, '') // 移除前綴 "台灣" 和郵遞區號
    .replace(/\s+/g, '')             // 移除所有空格
    .replace(/,/g, '')               // 移除所有逗號
    .trim();                         // 去掉首尾空格
}

//濾網
const highlightedProducts = [ //14
  "C200-Carbon前置濾網(盒)",
  "C200-Combo主濾網(盒)",
  "C200-Odors主濾網(盒)",
  "C200-Bio前置濾網(盒)",
  "C260-Bio前置濾網(盒)",
  "C260-Odors主濾網(盒)",
  "C360-Carbon前置濾網(盒)",
  "C360-Bio前置濾網(盒)",
  "C360-Odors主濾網(盒)",
  "C600-Carbon前置濾網(盒)",
  "C600-Combo主濾網(盒)",
  "C600-Odors主濾網(盒)",
  "C600-Bio前置濾網(盒)",
  "M1濾網(一盒2片)"
];
//設備
const highlightedProducts2 = [ //15
  "L1",
  "S1(白色)",
  "C200",
  "C260",
  "C360",
  "C600",
  "Smini主機",
  "Snano主機-黑",
  "Snano主機-紅",
  "Snano主機-綠",
  "Snano主機-藍",
  "S1 Pro",
  "L10",
  "A1",
  "M1"
];
//配件
const highlightedProducts3 = [ //21
  "M1車袋",
  "Smini Adaptor-US",
  "Smini專用桌架",
  "Smini皮革提帶",
  "Smini專用壁掛架",
  "Smini行動電池盒",
  "LG18650充電電池(組)",
  "E27節律球泡燈",
  "節律照明控制器",
  "LED驅動器(60W無風扇)(LPV-60-24)",
  "IS-08 P2 燈具藍芽控制器",
  "LED照明電源驅動器(40W 20V)",
  "HCL控制器",
  "掛板 (P2016 MOUNTING PLATE)",
  "LED電源供應器15V/3.4A (LRS-50-15)",
  "Snano磁吸支架",
  "Snano擴香片補充包",
  "Snano精油-健康呼吸",
  "Snano精油-樂活無鬱",
  "Snano精油-安神好眠",
  "Snano精油-清新健腦"
];

//商品統計按鈕
const ProductToExcel = () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Product');

  const header = [`本次轉換運單數量: ${waybillCount.value}`, "產品名稱", "數量"];
  worksheet.columns = [
    { width: 30 }, { width: 40 }, { width: 10 }
  ];

  const headerRow = worksheet.addRow(header);
  headerRow.height = 30; // 設定行高為 30
  //顏色
  headerRow.getCell(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F0F0F0' } // 設置淺灰色
  };
  headerRow.getCell(2).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F0F0F0' } // 設置淺灰色
  };
  headerRow.getCell(3).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F0F0F0' } // 設置淺灰色
  };
  headerRow.eachCell((cell) => {
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.font = { bold: true };
  });
  // 設置下邊框
  headerRow.getCell(1).border = {
    bottom: { style: 'thick', color: { argb: '000000' } }
  }; 
  headerRow.getCell(2).border = {
    bottom: { style: 'thick', color: { argb: '000000' } }
  }; 
  headerRow.getCell(3).border = {
    bottom: { style: 'thick', color: { argb: '000000' } }
  }; 

  const lines = outputText.value.split('\n');
  let isProductSection = false;
  let productData = [];

  lines.forEach((line) => {
    if (line.includes("【Total Product】")) {
      isProductSection = true;
      return;
    }
    if (isProductSection && line.trim() !== "") {
      const parts = line.split(" : ");
      if (parts.length === 2) {
        const product = parts[0].trim();
        const quantity = Number(parts[1].trim());
        productData.push([product, quantity]);
      }
    }
  });

  let category1 = 0; // 濾網
  let category2 = 0; //設備
  let category3 = 0; //配件
  let category4 = 0; //其他
  let totalQuantity = 0; // 記錄總數量

  productData.forEach(([product, quantity]) => {
    const trimmedProduct = product.trim(); // Trim spaces
    let category = ""; // 儲存分類名稱

    // 判斷產品屬於哪個分類
    if (highlightedProducts.includes(trimmedProduct)) {
      category1++;
      if (category1 === 1) {
        category = "濾網➜";
      }
    } else if (highlightedProducts2.includes(trimmedProduct)) {
      category2++;
      if (category2 === 1) {
        category = "設備➜";
      }
    } else if (highlightedProducts3.includes(trimmedProduct)) {
      category3++;
      if (category3 === 1) {
        category = "配件➜";
      }
    } else{
      category4++;
      if (category4 === 1) {
        category = "其他➜";
      }
    }

    // 記錄總數量
    totalQuantity += quantity;

    const dataRow = worksheet.addRow([category, trimmedProduct, quantity]);
    dataRow.height = 30;

    // 設置居中
    dataRow.getCell(1).alignment = {
      horizontal: 'center',
      vertical: 'middle', // 垂直居中
    };
    dataRow.getCell(2).alignment = {
      vertical: 'middle', // 垂直居中
      indent: 1 // 增加一點距離
    };
    dataRow.getCell(3).alignment = {
      horizontal: 'center', // 水平居中
      vertical: 'middle' // 垂直居中
    };

    // 字體
    dataRow.getCell(1).font = {
      bold: true,       // 粗體
      color: { argb: 'FF0000' } // 紅色
    };
    // 設置顏色和填充
    if (highlightedProducts.includes(trimmedProduct)) {
      dataRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F0F0F0' } // 設置淺灰色
      };
      dataRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DFFFDF' } // 設置淺綠色
      };
      dataRow.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DFFFDF' } // 設置淺綠色
      };
    } else if (highlightedProducts2.includes(trimmedProduct)) {
      dataRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F0F0F0' } // 設置淺灰色
      };
      dataRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D2E9FF' } // 設置淺藍色
      };
      dataRow.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D2E9FF' } // 設置淺藍色
      };
    } else if (highlightedProducts3.includes(trimmedProduct)) {
      dataRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F0F0F0' } // 設置淺灰色
      };
      dataRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D8D8EB' } // 設置淺紫色
      };
      dataRow.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D8D8EB' } // 設置淺紫色
      };
    } else{
      dataRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F0F0F0' } // 設置淺灰色
      };
      dataRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE4CA' } // 設置淺橘色
      };
      dataRow.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE4CA' } // 設置淺橘色
      };
    }
  });

  // footer
  const totalRow = worksheet.addRow(["總數➜", "", Number(totalQuantity)]);
  totalRow.height = 30; // 設定行高為 30
  // 字體
  totalRow.getCell(1).font = {
    bold: true,       // 粗體
    color: { argb: '000000' } 
  };
  totalRow.getCell(1).alignment = {
    horizontal: 'center', // 水平居中
    vertical: 'middle' // 垂直居中
  };
  totalRow.getCell(3).font = {
    bold: true,       // 粗體
  };
  // 設置上邊框
  totalRow.getCell(1).border = {
    top: { style: 'thick', color: { argb: '000000' } }
  }; 
  totalRow.getCell(2).border = {
    top: { style: 'thick', color: { argb: '000000' } }
  }; 
  totalRow.getCell(3).border = {
    top: { style: 'thick', color: { argb: '000000' } }
  }; 
  //對齊
  totalRow.getCell(3).alignment = {
    horizontal: 'center', // 水平居中
    vertical: 'middle' // 垂直居中
  };
  //顏色
  totalRow.getCell(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F0F0F0' } // 設置淺灰色
  };
  totalRow.getCell(2).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F0F0F0' } // 設置淺灰色
  };
  totalRow.getCell(3).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F0F0F0' } // 設置淺灰色
  };
  const today = new Date();
  const dateString = today.toISOString().slice(0, 10).replace(/-/g, '');

  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${dateString}_product_template.xlsx`;
    link.click();
  });

  
};

//運單彙總按鈕
const OrdersToExcel = () => {
  const lines = inputText.value.trim().split('\n');
  const orderMap = new Map();

  // 遍歷每一行數據並構建訂單
  lines.forEach((line) => {
    const parts = line.trim().split(/\t/);
    const address = normalizeAddress(parts[5]); 
    const [orderNumber, , , name, rawPhone, , , , , product, quantity] = parts;
    const phone = rawPhone.replace(/\s+/g, ''); // 移除所有空格

    const key = `${orderNumber}-${name}-${phone}-${address}`;

    if (!orderMap.has(key)) {
      orderMap.set(key, { orderNumber, name, phone, address, products: new Map() });
    }

    const order = orderMap.get(key);
    const currentQuantity = parseInt(quantity, 10) || 0;
    
    if (!order.products.has(product)) {
      order.products.set(product, 0);
    }

    order.products.set(product, order.products.get(product) + currentQuantity);
  });

  // 定義表頭
  const header = [
    "*客戶訂單號", "*收件方姓名", "收件方電話", "*收件方詳細地址", "*商品名稱", "*商品數量", "包裹備註", "代收貨款", "月結帳號", "附加服務內容"
  ];

  // 構建數據行
  const data = Array.from(orderMap.values()).map(({ orderNumber, name, phone, address, products }) => {
    const productDetails = Array.from(products.entries())
      .map(([product, totalQuantity]) => `${product}*${totalQuantity}`) // 保留商品名稱與原始數量
      .join('_');
    return [
      orderNumber,  // 客戶訂單號
      name,         // 收件方姓名
      phone,        // 收件方電話
      address,      // 收件方詳細地址
      productDetails, // 商品名稱 (包含原始數量)
      1,            // 商品數量固定為 1
      '', '', '', '' // 其他列為空
    ];
  });

  // 創建一個新的工作簿
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Orders');

  // 設置列寬
  worksheet.columns = [
    { width: 22 }, { width: 15 }, { width: 15 }, { width: 50 }, { width: 40 }, { width: 13 },
    { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }
  ];

  // 添加表頭
  worksheet.addRow(header).eachCell((cell, colNumber) => {
    // 設置統一的字體：微軟雅黑
    cell.font = { name: 'Microsoft YaHei', bold: true, size: 10 };

    // 第一欄
    if (colNumber === 1) {
      cell.font.color = { argb: "FF0000" }; // 紅字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FFF1E1" } }; // 淡米色背景
    }
    // 第二欄
    else if (colNumber === 2) {
      cell.font.color = { argb: "FF0000" }; // 紅字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FAD0D0" } }; // 淡红色背景
    }
    // 第三欄
    else if (colNumber === 3) {
      cell.font.color = { argb: "000000" }; // 黑字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FAD0D0" } }; // 淡紅色背景
    }
    // 第四欄
    else if (colNumber === 4) {
      cell.font.color = { argb: "FF0000" }; // 紅字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FAD0D0" } }; // 淡紅色背景
    }
    // 第五欄
    else if (colNumber === 5) {
      cell.font.color = { argb: "FF0000" }; // 紅字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FFF1E1" } }; // 淡米色背景
    }
    // 第六欄
    else if (colNumber === 6) {
      cell.font.color = { argb: "FF0000" }; // 紅字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FFF1E1" } }; // 淡米色背景
    }
    // 第七欄
    else if (colNumber === 7) {
      cell.font.color = { argb: "FF0000" }; // 紅字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "D9E4F4" } }; // 淡藍色背景
    }
    // 第八、九、十欄
    else if (colNumber >= 8 && colNumber <= 10) {
      cell.font.color = { argb: "000000" }; // 黑字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "D9EAD3" } }; // 淡綠色背景
    }

    // 其他統一設置
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });

  // 設置表頭行的高度為 30
  worksheet.getRow(1).height = 30;


  // 隱藏第 2 和第 3 列
  const row = worksheet.addRow([
    "必填，支持輸入字母和數字，max=64", 
    "必填，max=100", 
    "收件方手機號和固定電話至少填寫一項，max=20", 
    "必填，max=200", 
    "必填，max=100", 
    "必填，僅支持輸入數字，max=17", 
    "",
    "COD-代收貨款，非必選", 
    "若附加服務1選擇了COD，則卡號必填，max = 30", 
    "max = 30"
  ]);
  row.eachCell((cell) => {
    cell.alignment = { vertical: 'middle', wrapText: true }; 
  });
    
  const row2 = worksheet.addRow([
    "test2017062816", 
    "Li si1", 
    "178888****", 
    "777 Henderson Blvd, South Bay,1B,Folcroft, PA", 
    "Milk", 
    "", 
    "", 
    "", 
    ""
  ]);
  row2.eachCell((cell, colNumber) => {
    if (colNumber <= 3) {
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    } else {
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
  });

  worksheet.getRow(2).hidden = true;
  worksheet.getRow(3).hidden = true;

  // 添加數據行
  data.forEach((row) => {
    const newRow = worksheet.addRow(row);

    // 設置每行的高度
    newRow.height = 50; // 可以根據需要調整行高

    newRow.eachCell((cell, colNumber) => {
      cell.font = { size: 10 };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

      if (colNumber === 4 || colNumber === 5) {
        // d 列和 e 列水平靠左
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      } else {
        // 其他列水平居中
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
    });
  });



  // 獲取當前日期（格式：yyyyMMdd）
  const today = new Date();
  const dateString = today.toISOString().slice(0, 10).replace(/-/g, '');

  // 導出 Excel 文件，文件名為當前日期_Orders_template.xlsx
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${dateString}_orders_template.xlsx`;
    link.click();
  });
};

// MergeOrders按鈕
const showModal = ref(false); // 控制 modal 的显示
const modalRef = ref(null); // 引用 modal 元素
const handleOrderMerge = async() => {
  if (!inputText.value.trim()) {
    outputText.value = ''; // 清空輸出框
    showModal.value = false;
    showNotificationWithDelay('Please enter content...', 'error');
    return;
  }
  const lines = inputText.value.trim().split('\n'); // 分割每一行

  // 验证每行格式
  const invalidLines = lines.map((line, index) => {
    const validationError = validateLine(line); // 使用 validateLine 函數驗證每行
    outputText.value = '';
    showModal.value = false;
    return validationError ? `Line ${index + 1}: ${validationError}` : null;
  }).filter(error => error !== null); // 過濾出有錯誤的行

  if (invalidLines.length > 0) {
    outputText.value = '';
    showModal.value = false;
    showNotificationWithDelay(invalidLines.join('\n'), 'error'); // 傳遞換行錯誤訊息
    return;
  }

  const orderMap = new Map(); // 儲存訂單數據
  const productSummary = new Map(); // 統計所有商品總數量

  lines.forEach((line) => {
    const parts = line.trim().split(/\t/); // 以 Tab 分隔每行
    const orderNumber = parts[0]; // 訂單號
    const name = parts[3]; // 姓名
    const phone = parts[4]; // 電話
    const address = normalizeAddress(parts[5]); // 地址
    const product = parts[9]; // 商品名稱
    const quantity = parseInt(parts[10], 10); // 商品數量

    const key = `${orderNumber}-${name}-${phone}-${address}`;

    if (!orderMap.has(key)) {
      orderMap.set(key, {
        orderNumber,
        name,
        phone,
        address,
        products: new Map(),
      });
    }

    const order = orderMap.get(key);

    if (!order.products.has(product)) {
      order.products.set(product, 0);
    }

    order.products.set(product, order.products.get(product) + quantity);

    // 統計商品總數
    if (!productSummary.has(product)) {
      productSummary.set(product, 0);
    }
    productSummary.set(product, productSummary.get(product) + quantity);
  });

  // 計算合併後的運單數量
  waybillCount.value = orderMap.size;

  // 獲取商品統計並轉換為數組
  const productSummaryOutput = Array.from(productSummary.entries());

  // 排序規則
  const sortedProductSummary = productSummaryOutput.sort((a, b) => {
    const indexA = highlightedProducts.indexOf(a[0]);
    const indexB = highlightedProducts.indexOf(b[0]);

    const indexA2 = highlightedProducts2.indexOf(a[0]);
    const indexB2 = highlightedProducts2.indexOf(b[0]);

    const indexA3 = highlightedProducts3.indexOf(a[0]);
    const indexB3 = highlightedProducts3.indexOf(b[0]);

    if (indexA !== -1 && indexB !== -1) return indexA - indexB; // `highlightedProducts` 按原數組順序
    if (indexA !== -1) return -1; // `highlightedProducts` 優先
    if (indexB !== -1) return 1;

    if (indexA2 !== -1 && indexB2 !== -1) return indexA2 - indexB2; // `highlightedProducts2` 按原數組順序
    if (indexA2 !== -1) return -1; // `highlightedProducts2` 次優先
    if (indexB2 !== -1) return 1;

    if (indexA3 !== -1 && indexB3 !== -1) return indexA3 - indexB3; // `highlightedProducts3` 按原數組順序
    if (indexA3 !== -1) return -1; // `highlightedProducts3` 次次優先
    if (indexB3 !== -1) return 1;

    return a[0].localeCompare(b[0]); // 其他商品按名稱排序
  });



  // 格式化輸出
  const formattedProductSummary = sortedProductSummary
    .map(([product, totalQuantity]) => `${product} : ${totalQuantity}`)
    .join('\n');

  // 組合最終輸出
  const finalOutput = `【Total Orders】\n${waybillCount.value} Unit\n\n【Total Product】\n${formattedProductSummary}`;

  outputText.value = finalOutput;
  showNotificationWithDelay('Order Merged Successfully！', 'success');
  // 彈出 modal
  if (finalOutput !== '') {
    showModal.value = true; // 打開 modal
    await nextTick(); // 等待 DOM 更新
    const modal = new bootstrap.Modal(modalRef.value); // 使用 Bootstrap 的 modal
    modal.show(); // 顯示 modal
  }
};

// Clear按鈕
const handleClear = () => {
  inputText.value = '';
  outputText.value = '';
  showModal.value = false;
};



</script>

<template>
  <div class="container">
    <!-- 動態通知 -->
    <div
      class="notification"
      :class="[notificationType === 'success' ? 'notification-success' : 'notification-error', showNotification ? 'notification-show' : '']"
    >
      <div v-html="notificationMessage"></div> <!-- 使用 v-html 来渲染带 HTML 的内容 -->
    </div>
    
    <!-- 輸入與輸出區 -->
    <div class="row mt-5">
      <div class="col-12 d-flex justify-content-center">
        <textarea
          class="form-control"
          v-model="inputText"
          @input="handleInput"
          placeholder="Enter orders (one per line)..."
        ></textarea>
      </div>
    </div>

    <!-- 操作按钮 -->
    <div class="row">
      <div class="col-12 d-flex justify-content-center mt-5">
        <button class="convert" @click="handleOrderMerge">Merge Orders</button>
        <button class="cls ms-3" @click="handleClear">Clear</button>
      </div>
    </div>
  </div>

  <!-- Modal -->
  <div  v-if="showModal" class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true" ref="modalRef">
    <div class="modal-dialog d-flex align-items-center">
      <div class="modal-content">
        <div class="modal-header d-flex justify-content-center">
          <h5 class="modal-title " id="exampleModalLabel">Export to Excel</h5>
        </div>
        <div class="modal-body d-flex justify-content-between">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal" @click="ProductToExcel">Product Statistics</button>
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal" @click="OrdersToExcel">Orders</button>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
textarea {
  width: 100%; 
  resize: none;
  height: 400px;
  border: 2px solid #000;
}
@media (max-width: 767px) {
  textarea, .output {
    height: 250px; 
  }
}
textarea:focus {
  outline: none;
  border-color: black;
}
.output {
  width: 100%; 
  border: 2px solid #000;
  height: 400px;
  border-radius: 5px;
  padding: 10px;
  white-space: pre-line;
  overflow-y: auto;
}
@media (max-width: 767px) {
  .output {
    height: 250px; 
  }
}
.convert,
.cls {
  border-radius: 50px;
  padding: 10px 20px;
  color: white;
  font-size: 16px;
  border: none;
  cursor: pointer;
  transition: background-color 0.3s ease;
  width: 100%; 
  max-width: 200px; 
}
.convert {
  background-color: #4caf50;
}
.convert:hover {
  background-color: #45a049;
}
.cls {
  background-color: #ce0000;
}
.cls:hover {
  background-color: #ae0000;
}
.notification {
  position: fixed;
  bottom: 50px;
  right: -300px;
  width: 350px;
  padding: 15px;
  border-radius: 5px;
  font-size: 16px;
  transition: right 0.3s ease-in-out, opacity 0.3s;
  opacity: 0;
}
.notification-show {
  right: 20px;
  opacity: 1;
}
.notification-success {
  background-color: #d4edda;
  border: 1px solid #c3e6cb;
  color: #155724;
}
.notification-error {
  background-color: #f8d7da;
  border: 1px solid #f5c6cb;
  color: #721c24;
}
.btn{
  color: white;
  width: 45%;
  padding: 20px;
}
.modal {
  background-color: rgba(0, 0, 0, 0.2);
}
.modal-dialog{
  margin-top: 200px;
}
.modal-body{
  padding: 30px;
}
.clickable-button:disabled {
  pointer-events: auto; /* 允許點擊事件 */
}
</style>
