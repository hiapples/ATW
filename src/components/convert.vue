<script setup>
import { ref } from 'vue';

const inputText = ref(''); // 輸入框內容
const outputText = ref(''); // 整合輸出框
const notificationMessage = ref(''); // 通知訊息
const notificationType = ref(''); // 通知類型
const showNotification = ref(false); // 控制通知顯示
let notificationTimeout; // 計時器參考

// 顯示通知
const showNotificationWithDelay = (message, type) => {
  if (notificationTimeout) clearTimeout(notificationTimeout);
  showNotification.value = false;

  setTimeout(() => {
    notificationMessage.value = message;
    notificationType.value = type;
    showNotification.value = true;
    notificationTimeout = setTimeout(() => {
      showNotification.value = false;
    }, 3000);
  }, 300);
};

// 檢查是否包含必要項目
const validateLine = (line) => {
  const parts = line.split(/\s+/); // 支援多個空白或 Tab 分隔
  return parts.length >= 4; // 確保至少有四部分：收件人、電話、地址、品項
};

// 整合訂單邏輯
const handleOrderMerge = () => {
  if (!inputText.value.trim()) {
    outputText.value = ''; // 清空輸出框
    showNotificationWithDelay('Please enter content...', 'error');
    return;
  }

  const lines = inputText.value.trim().split('\n'); // 分割每一行

  // 驗證每行格式
  const invalidLines = lines.filter((line) => !validateLine(line.trim()));
  if (invalidLines.length > 0) {
    showNotificationWithDelay(`Error: Invalid lines detected:\n${invalidLines.join('\n')}`, 'error');
    return;
  }

  const orderMap = new Map(); // 儲存訂單資料

  lines.forEach((line) => {
    const parts = line.trim().split(/\s+/); // 分割每行
    const key = parts.slice(0, 3).join(' '); // 客戶資訊作為鍵
    const product = parts[3]; // 商品類型

    if (!orderMap.has(key)) {
      orderMap.set(key, {});
    }

    const products = orderMap.get(key);
    products[product] = (products[product] || 0) + 1; // 累加商品數量
  });

  // 整合輸出結果
  const output = Array.from(orderMap.entries())
    .map(([key, products]) => {
      const productSummary = Object.entries(products)
        .map(([product, count]) => `${product}*${count}`) // 商品部分用 * 數量標示
        .join('_'); // 商品部分用 _ 連接
      return `${key}    ${productSummary}__`; // 每行用 4 個空白鍵分隔，最後附加 __
    })
    .join('\n');

  outputText.value = output; // 更新輸出框
  showNotificationWithDelay('Order Merged Successfully!', 'success');
};


// 清除內容
const handleClear = () => {
  inputText.value = '';
  outputText.value = '';
};
</script>

<template>
  <div class="container">
    <!-- 動態通知 -->
    <div
      class="notification"
      :class="[notificationType === 'success' ? 'notification-success' : 'notification-error', showNotification ? 'notification-show' : '']"
    >
      {{ notificationMessage }}
    </div>

    <!-- 輸入與輸出區 -->
    <div class="row mt-5">
      <div class="col d-flex justify-content-center">
        <textarea
          class="form-control"
          v-model="inputText"
          placeholder="Enter orders (one per line)..."
        ></textarea>
      </div>
      <div class="col d-flex justify-content-center">
        <div class="output">{{ outputText }}</div> <!-- 顯示 outputText -->
      </div>
    </div>

    <!-- 操作按鈕 -->
    <div class="row">
      <div class="col mt-5 d-flex justify-content-center">
        <button class="convert" @click="handleOrderMerge">Merge Orders</button>
        <button class="cls ms-3" @click="handleClear">Clear</button>
      </div>
    </div>
  </div>
</template>

<style scoped>
textarea {
  width: 95%;
  resize: none;
  height: 400px;
  border: 2px solid #000;
}
textarea:focus {
  outline: none;
  border-color: black;
}
.output {
  width: 95%;
  border: 2px solid #000;
  height: 400px;
  border-radius: 5px;
  padding: 10px;
  white-space: pre-line;
  overflow-y: auto;
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
  width: 10%;
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
/* 動態通知樣式 */
.notification {
  position: fixed;
  bottom: 50px;
  right: -300px;
  width: 250px;
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
