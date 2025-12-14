// config.js

const API_CONFIG = {
  // 正式環境的 API URL
  apiUrl: "https://script.google.com/macros/s/AKfycbwip-XxLPYl02S0VPj_5Xi9lFpW59yQ9Hr419RvJmOGqpMxy3hQqYraPWfecq99Qc4V/exec",
  
  // 新增回呼網址
  redirectUrl: "https://eric693.github.io/check_manager_plus/"
  // 你也可以在這裡加入其他設定，例如：
  // timeout: 5000,
  // version: 'v4.2.3'
};
// 👇 新增：為了兼容性，同時定義全域變數 apiUrl
const apiUrl = API_CONFIG.apiUrl;
