<script setup>
import { ref } from 'vue';

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
    notificationMessage.value = message;
    notificationType.value = type;
    showNotification.value = true;
    notificationTimeout = setTimeout(() => {
      showNotification.value = false;
    }, 3000);
  }, 300);
};

// 检查是否包含必要项
const validateLine = (line) => {
  const parts = line.split(/\t/); // 以 Tab 分隔
  return parts.length >= 11; // 确保至少有 11 部分
};

// 合并订单逻辑
const handleOrderMerge = () => {
  if (!inputText.value.trim()) {
    outputText.value = ''; // 清空输出框
    showNotificationWithDelay('Please enter content...', 'error');
    return;
  }

  const lines = inputText.value.trim().split('\n'); // 分割每一行

  // 验证每行格式
  const invalidLines = lines.filter((line) => !validateLine(line.trim()));
  if (invalidLines.length > 0) {
    showNotificationWithDelay(`Error: Invalid lines detected:\n${invalidLines.join('\n')}`, 'error');
    return;
  }

  const orderMap = new Map(); // 存储订单数据

  lines.forEach((line) => {
    const parts = line.trim().split(/\t/); // 以 Tab 分隔每行
    const orderNumber = parts[0]; // 订单号
    const name = parts[3]; // 姓名
    const phone = parts[4]; // 电话
    const address = parts[5]; // 地址
    const product = parts[9]; // 商品名称
    const quantity = parseInt(parts[10], 10); // 商品数量

    // 生成一个唯一的键，判断是否存在相同的订单（通过订单号、姓名、电话、地址来判断）
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
  });

  // 合并输出结果
  const output = Array.from(orderMap.values())
    .map(({ orderNumber, name, phone, address, products }) => {
      const productDetails = Array.from(products.entries())
        .map(([product, totalQuantity]) => `${product}*${totalQuantity}`)
        .join('_');
      return `${orderNumber}\t${name}  ${phone}\t${address}\t${productDetails}__`;
    })
    .join('\n');

  outputText.value = output; // 更新输出框
  showNotificationWithDelay('Order Merged Successfully！', 'success');
};

// 清空内容
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
      <div class="col-12 col-md-6 d-flex justify-content-center">
        <textarea
          class="form-control"
          v-model="inputText"
          placeholder="Enter orders (one per line)..."
        ></textarea>
      </div>
      <div class="col-12 mt-2 mt-md-0 col-md-6 d-flex justify-content-center">
        <div class="output">{{ outputText }}</div> <!-- 顯示 outputText -->
      </div>
    </div>

    <!-- 操作按鈕 -->
    <div class="row">
      <div class="col-12 d-flex justify-content-center mt-5">
        <button class="convert" @click="handleOrderMerge">Merge Orders</button>
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
.cls {
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
