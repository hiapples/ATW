<script setup>
import { ref } from 'vue';
import ExcelJS from 'exceljs';

const inputText = ref(''); // 输入框内容
const outputText = ref(''); // 输出结果
const waybillCount = ref(''); //運單數量
const notificationMessage = ref(''); // 通知消息
const notificationType = ref(''); // 通知类型
const showNotification = ref(false); // 控制通知显示
let notificationTimeout; // 计时器参考

// 显示通知
const showNotificationWithDelay = (message, type) => {
  if (notificationTimeout) clearTimeout(notificationTimeout);
  showNotification.value = false;

  setTimeout(() => {
    notificationMessage.value = message.replace(/\n/g, '<br>'); // 替换换行符为 <br> 标签
    notificationType.value = type;
    showNotification.value = true;
    notificationTimeout = setTimeout(() => {
      showNotification.value = false;
    }, 5000); // 显示通知 5 秒后关闭
  }, 300);
};

// 固定列数检查
const validateLine = (line) => {
  const parts = line.split(/\t+/); // 按 Tab 分隔

  if (parts.length < 11) {
    outputText.value = '';
    return 'Error: Too few columns (less than 11)';
  }

  const orderNumber = parts[0]; // 銷貨單號
  const orderNumberRegex = /^[a-zA-Z0-9#]*$/; // 正则表达式，要求只包含英文和数字

  if (!orderNumberRegex.test(orderNumber)) {
    outputText.value = '';
    return 'Error: Order number must contain only letters or digits';
  }

  const phone = parts[4]; // "电话号码"在第 5 列（索引为 4）
  const phoneRegex = /^\d+$/; // 正则表达式，要求只包含数字

  if (!phoneRegex.test(phone)) {
    outputText.value = '';
    return 'Error: Phone number must contain only digits';
  }

  const quantity = parts[10]; // "數量" 在第 11 列（索引 10）
  const quantityRegex = /^[1-9]\d*$/; // 正则表达式，要求只包含数字

  if (!quantityRegex.test(quantity)) {
    outputText.value = '';
    return 'Error: Quantity must be a number';
  }

  if (parts.length > 11) {
    outputText.value = '';
    return 'Error: Too many columns (greater than 11)';
  }


  return null; // 如果没有错误
};

//輸入處理
const handleInput = () => {
  // 使用正則表達式將兩個 Tab 中間加上空格
  inputText.value = inputText.value.replace(/\t\t/g, '\t \t');
  // 按换行符拆分输入内容为行
  let lines = inputText.value.split('\n');

  // 初始化结果数组
  let result = [];
  let tempLine = '';

  lines.forEach((line) => {
    // 去掉前后空格
    line = line.trim();

    // 检查行是否以数字结尾（即是完整记录）
    if (/.*\d$/.test(line)) {
      // 如果当前有临时行，合并到结果中
      if (tempLine) {
        tempLine += ' ' + line; // 合并记录
        result.push(tempLine.trim());
        tempLine = ''; // 清空临时行
      } else {
        // 如果没有临时行，直接加入结果
        result.push(line);
      }
    } else {
      // 行内数据需要合并到临时行
      tempLine += ' ' + line;
    }
  });

  // 如果最后有剩余的临时行，加入结果
  if (tempLine) {
    result.push(tempLine.trim());
  }

  // 将结果数组合并为字符串，保留记录间的换行符
  inputText.value = result.join('\n');
};

//地址標準化
function normalizeAddress(address) {
  return address
    .replace(/^台灣\s*\d{3}\s*/, '') // 移除前缀 "台灣" 和邮递区号
    .replace(/\s+/g, '')             // 移除所有空格
    .replace(/,/g, '')               // 移除所有逗号
    .trim();                         // 去掉首尾空格
}

// 合并订单逻辑
const handleOrderMerge = () => {
  if (!inputText.value.trim()) {
    outputText.value = ''; // 清空输出框
    showNotificationWithDelay('Please enter content...', 'error');
    return;
  }
  isModalVisible.value = true;
  const lines = inputText.value.trim().split('\n'); // 分割每一行

  // 验证每行格式
  const invalidLines = lines.map((line, index) => {
    const validationError = validateLine(line); // 使用 validateLine 函数验证每行
    outputText.value = '';
    return validationError ? `Line ${index + 1}: ${validationError}` : null;
  }).filter(error => error !== null); // 过滤出有错误的行

  if (invalidLines.length > 0) {
    outputText.value = '';
    showNotificationWithDelay(invalidLines.join('\n'), 'error'); // 传递换行错误信息
    return;
  }

  const orderMap = new Map(); // 存储订单数据
  const productSummary = new Map(); // 统计所有商品总数量

  lines.forEach((line) => {
    const parts = line.trim().split(/\t/); // 以 Tab 分隔每行
    const orderNumber = parts[0]; // 订单号
    const name = parts[3]; // 姓名
    const phone = parts[4]; // 电话
    const address = normalizeAddress(parts[5]); // 地址
    const product = parts[9]; // 商品名称
    const quantity = parseInt(parts[10], 10); // 商品数量

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

    // 统计商品总数
    if (!productSummary.has(product)) {
      productSummary.set(product, 0);
    }
    productSummary.set(product, productSummary.get(product) + quantity);
  });

  // 计算合并后的运单数量
   waybillCount.value = orderMap.size;

  // 获取商品统计并转换为数组
  const productSummaryOutput = Array.from(productSummary.entries());

  // 排序规则
  const sortedProductSummary = productSummaryOutput.sort((a, b) => {
    const indexA = highlightedProducts.indexOf(a[0]);
    const indexB = highlightedProducts.indexOf(b[0]);

    const indexA2 = highlightedProducts2.indexOf(a[0]);
    const indexB2 = highlightedProducts2.indexOf(b[0]);

    if (indexA !== -1 && indexB !== -1) return indexA - indexB; // `highlightedProducts` 按原数组顺序
    if (indexA !== -1) return -1; // `highlightedProducts` 优先

    if (indexB !== -1) return 1; // `highlightedProducts` 在 `highlightedProducts2` 之前

    if (indexA2 !== -1 && indexB2 !== -1) return indexA2 - indexB2; // `highlightedProducts2` 按原数组顺序
    if (indexA2 !== -1) return -1; // `highlightedProducts2` 次优先

    if (indexB2 !== -1) return 1; // `highlightedProducts2` 在其他商品之前

    return a[0].localeCompare(b[0]); // 其他商品按名称排序
  });


  // 格式化输出
  const formattedProductSummary = sortedProductSummary
    .map(([product, totalQuantity]) => `${product} : ${totalQuantity}`)
    .join('\n');

  // 组合最终输出
  const finalOutput = `【Total Orders】\n${waybillCount.value} Unit\n\n【Total Product】\n${formattedProductSummary}`;

  outputText.value = finalOutput;
  showNotificationWithDelay('Order Merged Successfully！', 'success');
};


const highlightedProducts = [ //14
  "C200-Carbon前置濾網(盒)",
  "C200-Combo主濾網(盒)",
  "C200-Odors主濾網(盒)",
  "C200-Bio前置濾網(盒)",
  "C600-Carbon前置濾網(盒)",
  "C600-Combo主濾網(盒)",
  "C600-Odors主濾網(盒)",
  "C600-Bio前置濾網(盒)",
  "C360-Carbon前置濾網(盒)",
  "C360-Bio前置濾網(盒)",
  "C360-Odors主濾網(盒)",
  "C260-Bio前置濾網(盒)",
  "C260-Odors主濾網(盒)",
  "M1濾網(一盒2片)"
];
const highlightedProducts2 = [ //15
  "A1",
  "M1",
  "S1(白色)",
  "C200",
  "C600",
  "C360",
  "C260",
  "S1 Pro",
  "Smini主機",
  "L10",
  "L1",
  "Snano主機-黑",
  "Snano主機-紅",
  "Snano主機-綠",
  "Snano主機-藍"
];
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

  let category1 = 0; // 紀錄次數
  let category2 = 0;
  let category3 = 0;
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




const OrdersToExcel = () => {
  const lines = inputText.value.trim().split('\n');
  const orderMap = new Map();

  // 遍历每一行数据并构建订单
  lines.forEach((line) => {
    const parts = line.trim().split(/\t/);
    const address = normalizeAddress(parts[5]); // 假设 `address` 在 `parts` 数组中的第 6 个位置（索引为 5）
    const [orderNumber, , , name, phone, , , , , product, quantity] = parts;

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

  // 定义表头
  const header = [
    "*客戶訂單號", "*收件方姓名", "收件方電話", "*收件方詳細地址", "*商品名稱", "*商品數量", "包裹備註", "代收貨款", "月結帳號", "附加服務內容"
  ];

  // 构建数据行
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
      '', '', '', '' // 其他列为空
    ];
  });

  // 创建一个新的工作簿
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Orders');

  // 设置列宽
  worksheet.columns = [
    { width: 22 }, { width: 15 }, { width: 15 }, { width: 50 }, { width: 40 }, { width: 13 },
    { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }
  ];

  // 添加表头
  worksheet.addRow(header).eachCell((cell, colNumber) => {
    // 设置统一的字体：微软雅黑
    cell.font = { name: 'Microsoft YaHei', bold: true, size: 10 };

    // 第一栏
    if (colNumber === 1) {
      cell.font.color = { argb: "FF0000" }; // 红字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FFF1E1" } }; // 淡米色背景
    }
    // 第二栏
    else if (colNumber === 2) {
      cell.font.color = { argb: "FF0000" }; // 红字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FAD0D0" } }; // 淡红色背景
    }
    // 第三栏
    else if (colNumber === 3) {
      cell.font.color = { argb: "000000" }; // 黑字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FAD0D0" } }; // 淡红色背景
    }
    // 第四栏
    else if (colNumber === 4) {
      cell.font.color = { argb: "FF0000" }; // 红字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FAD0D0" } }; // 淡红色背景
    }
    // 第五栏
    else if (colNumber === 5) {
      cell.font.color = { argb: "FF0000" }; // 红字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FFF1E1" } }; // 淡米色背景
    }
    // 第六栏
    else if (colNumber === 6) {
      cell.font.color = { argb: "FF0000" }; // 红字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "FFF1E1" } }; // 淡米色背景
    }
    // 第七栏
    else if (colNumber === 7) {
      cell.font.color = { argb: "FF0000" }; // 红字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "D9E4F4" } }; // 淡蓝色背景
    }
    // 第八、九、十栏
    else if (colNumber >= 8 && colNumber <= 10) {
      cell.font.color = { argb: "000000" }; // 黑字
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: "D9EAD3" } }; // 淡绿色背景
    }

    // 其他统一设置
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });

  // 设置表头行的高度为 30
  worksheet.getRow(1).height = 30;


  // 隐藏第 2 和第 3 列
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
    cell.alignment = { vertical: 'middle', wrapText: true }; // Align vertically and wrap text
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
      // Horizontally center the first three columns
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    } else {
      // For the remaining columns, keep the previous vertical alignment and wrapping
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
  });

  worksheet.getRow(2).hidden = true;
  worksheet.getRow(3).hidden = true;

  // 添加数据行
  data.forEach((row) => {
    const newRow = worksheet.addRow(row);

    // 设置每行的高度
    newRow.height = 50; // 可以根据需要调整行高

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



  // 获取当前日期（格式：yyyyMMdd）
  const today = new Date();
  const dateString = today.toISOString().slice(0, 10).replace(/-/g, '');

  // 导出 Excel 文件，文件名为当前日期_Orders_template.xlsx
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${dateString}_orders_template.xlsx`;
    link.click();
  });
};



// 控制modal彈出
const isModalVisible = ref(false);
// 打開 modal
const handleModal = () => {
  // 如果 outputText 為空，則不打開 modal
  if (outputText.value === '') {
    showNotificationWithDelay('The output box is empty, please merge the orders first.', 'error')
  } else {
    // 如果有值，則打開 modal
    isModalVisible.value = true;
  }
};

// 清空内容
const handleClear = () => {
  inputText.value = '';
  outputText.value = '';
  isModalVisible.value = false;
};
</script>

<template>
  <div class="container">
    <!-- 动态通知 -->
    <div
      class="notification"
      :class="[notificationType === 'success' ? 'notification-success' : 'notification-error', showNotification ? 'notification-show' : '']"
    >
      <div v-html="notificationMessage"></div> <!-- 使用 v-html 来渲染带 HTML 的内容 -->
    </div>
    
    <!-- 输入与输出区 -->
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
        <button class="export ms-3" @click="handleModal" data-bs-toggle="modal" data-bs-target="#exampleModal">Export to Excel</button> <!-- 导出按钮 -->
        <button class="cls ms-3" @click="handleClear">Clear</button>
      </div>
    </div>
  </div>

  <!-- Modal -->
  <div v-if="isModalVisible" class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
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
  width: 100%; /* Full width on small screens */
  resize: none;
  height: 400px;
  border: 2px solid #000;
}
@media (max-width: 767px) {
  textarea, .output {
    height: 250px; /* Half the height on small screens */
  }
}
textarea:focus {
  outline: none;
  border-color: black;
}
.output {
  width: 100%; /* Full width on small screens */
  border: 2px solid #000;
  height: 400px;
  border-radius: 5px;
  padding: 10px;
  white-space: pre-line;
  overflow-y: auto;
}
@media (max-width: 767px) {
  .output {
    height: 250px; /* Half the height on small screens */
  }
}
.convert,
.cls,
.export {
  border-radius: 50px;
  padding: 10px 20px;
  color: white;
  font-size: 16px;
  border: none;
  cursor: pointer;
  transition: background-color 0.3s ease;
  width: 100%; /* Full width on small screens */
  max-width: 200px; /* Max width for larger screens */
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
.export {
  background-color: #007bff;
}
.export:hover {
  background-color: #0056b3;
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
</style>
