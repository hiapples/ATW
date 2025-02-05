<script setup>
import { ref } from 'vue';
import ExcelJS from 'exceljs';

const inputText = ref(''); // 输入框内容
const outputText = ref(''); // 输出结果
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
  const waybillCount = orderMap.size;

  // 添加商品统计
  const productSummaryOutput = Array.from(productSummary.entries())
    .map(([product, totalQuantity]) => `${product} : ${totalQuantity}`)
    .join('\n');

  const finalOutput = `【運單總數】\n${waybillCount} 單\n\n【商品統計】\n${productSummaryOutput}`;

  outputText.value = finalOutput;
  showNotificationWithDelay('Order Merged Successfully！', 'success');
};



const exportToExcel = () => {
  // 检查输出框是否有内容
  if (!outputText.value.trim()) {
    showNotificationWithDelay('The output box is empty, please merge the orders first.', 'error');
    return;
  }

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

  const sheetData = [header, ...data];

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

// 清空内容
const handleClear = () => {
  inputText.value = '';
  outputText.value = '';
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
      <div class="col-12 col-md-6 d-flex justify-content-center">
        <textarea
          class="form-control"
          v-model="inputText"
          @input="handleInput"
          placeholder="Enter orders (one per line)..."
        ></textarea>
      </div>
      <div class="col-12 mt-2 mt-md-0 col-md-6 d-flex justify-content-center">
        <div class="output">{{ outputText }}</div> <!-- 显示 outputText -->
      </div>
    </div>

    <!-- 操作按钮 -->
    <div class="row">
      <div class="col-12 d-flex justify-content-center mt-5">
        <button class="convert" @click="handleOrderMerge">Merge Orders</button>
        <button class="export ms-3" @click="exportToExcel">Export to Excel</button> <!-- 导出按钮 -->
        <button class="cls ms-3" @click="handleClear">Clear</button>
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
</style>
