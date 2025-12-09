// LineNotification.gs - LINE 推播通知系統（完整版）

// ==================== 常數設定 ====================
const LINE_CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ACCESS_TOKEN");
const LINE_MESSAGING_API_URL = "https://api.line.me/v2/bot/message/push";

/**
 * 發送 LINE 推播訊息
 * @param {string} userId - LINE User ID
 * @param {Object} flexMessage - Flex Message 物件
 */
function sendLineNotification_(userId, flexMessage) {
  const payload = {
    to: userId,
    messages: [flexMessage]
  };
  
  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(LINE_MESSAGING_API_URL, options);
    const result = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200) {
      Logger.log(`✅ LINE 通知已發送給 ${userId}`);
      return { ok: true };
    } else {
      Logger.log(`❌ LINE 通知發送失敗: ${result.message}`);
      return { ok: false, error: result.message };
    }
  } catch (err) {
    Logger.log(`❌ LINE API 錯誤: ${err.message}`);
    return { ok: false, error: err.message };
  }
}

// ==================== Flex Message 模板 ====================

/**
 * 🔔 忘記打卡提醒
 */
function createForgotPunchNotification(employeeName, date, punchType) {
  return {
    type: "flex",
    altText: `⚠️ ${employeeName}，您忘記${punchType}打卡了！`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "⚠️ 忘記打卡提醒",
            weight: "bold",
            size: "xl",
            color: "#FF6B6B",
            align: "center"
          }
        ],
        backgroundColor: "#FFF5F5",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "日期",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: date,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "狀態",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `忘記${punchType}打卡`,
                    wrap: true,
                    color: "#FF6B6B",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              }
            ]
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "text",
            text: "請盡快進行補打卡，避免影響出勤記錄！",
            size: "sm",
            color: "#666666",
            margin: "lg",
            wrap: true
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "primary",
            height: "sm",
            action: {
              type: "uri",
              label: "立即補打卡",
              uri: "https://eric693.github.io/check_manager_plus/"
            },
            color: "#4CAF50"
          },
          {
            type: "button",
            style: "link",
            height: "sm",
            action: {
              type: "uri",
              label: "查看打卡記錄",
              uri: "https://eric693.github.io/check_manager_plus/"
            }
          }
        ],
        flex: 0
      }
    }
  };
}

/**
 * ✅ 補打卡審核通知（核准）
 */
function createPunchApprovedNotification(employeeName, date, time, punchType, reviewer) {
  return {
    type: "flex",
    altText: `✅ 您的補打卡申請已核准`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "✅ 審核通過",
            weight: "bold",
            size: "xl",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#4CAF50",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "text",
            text: "您的補打卡申請已通過審核",
            size: "sm",
            color: "#666666",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "日期",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: date,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "時間",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: time,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "類型",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `${punchType}打卡`,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "審核人",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reviewer,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              }
            ]
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "primary",
            height: "sm",
            action: {
              type: "uri",
              label: "查看詳情",
              uri: "https://eric693.github.io/check_manager_plus/"
            },
            color: "#4CAF50"
          }
        ],
        flex: 0
      }
    }
  };
}

/**
 * ❌ 補打卡審核通知（拒絕）
 */
function createPunchRejectedNotification(employeeName, date, time, punchType, reviewer, reason) {
  return {
    type: "flex",
    altText: `❌ 您的補打卡申請已被拒絕`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "❌ 審核未通過",
            weight: "bold",
            size: "xl",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#FF6B6B",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "text",
            text: "您的補打卡申請已被拒絕",
            size: "sm",
            color: "#666666",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "日期",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: date,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "時間",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: time,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "審核人",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reviewer,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              },
              {
                type: "separator",
                margin: "md"
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                margin: "md",
                contents: [
                  {
                    type: "text",
                    text: "拒絕原因",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reason || "未提供",
                    wrap: true,
                    color: "#FF6B6B",
                    size: "sm",
                    flex: 5
                  }
                ]
              }
            ]
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "link",
            height: "sm",
            action: {
              type: "uri",
              label: "查看詳情",
              uri: "https://eric693.github.io/check_manager_plus/"
            }
          }
        ],
        flex: 0
      }
    }
  };
}

/**
 * ✅ 請假審核通知（核准）
 */
function createLeaveApprovedNotification(employeeName, leaveType, startDate, endDate, days, reviewer) {
  return {
    type: "flex",
    altText: `✅ 您的${leaveType}申請已核准`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "✅ 請假核准通知",
            weight: "bold",
            size: "xl",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#2196F3",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "text",
            text: "您的請假申請已通過審核",
            size: "sm",
            color: "#666666",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "假別",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: leaveType,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "期間",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `${startDate} ~ ${endDate}`,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "天數",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `${days} 天`,
                    wrap: true,
                    color: "#2196F3",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "審核人",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reviewer,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              }
            ]
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "text",
            text: "🎉 祝您有個愉快的假期！",
            size: "sm",
            color: "#2196F3",
            margin: "lg",
            align: "center",
            weight: "bold"
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "primary",
            height: "sm",
            action: {
              type: "uri",
              label: "查看假期餘額",
              uri: "https://eric693.github.io/check_manager_plus/"
            },
            color: "#2196F3"
          }
        ],
        flex: 0
      }
    }
  };
}

/**
 * ❌ 請假審核通知（拒絕）
 */
function createLeaveRejectedNotification(employeeName, leaveType, startDate, endDate, days, reviewer, reason) {
  return {
    type: "flex",
    altText: `❌ 您的${leaveType}申請已被拒絕`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "❌ 請假未核准",
            weight: "bold",
            size: "xl",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#FF9800",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "text",
            text: "您的請假申請未通過審核",
            size: "sm",
            color: "#666666",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "假別",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: leaveType,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "期間",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `${startDate} ~ ${endDate}`,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "審核人",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reviewer,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              },
              {
                type: "separator",
                margin: "md"
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                margin: "md",
                contents: [
                  {
                    type: "text",
                    text: "拒絕原因",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reason || "未提供",
                    wrap: true,
                    color: "#FF9800",
                    size: "sm",
                    flex: 5
                  }
                ]
              }
            ]
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "link",
            height: "sm",
            action: {
              type: "uri",
              label: "重新申請",
              uri: "https://eric693.github.io/check_manager_plus/"
            }
          }
        ],
        flex: 0
      }
    }
  };
}

/**
 * ✅ 加班審核通知（核准）
 */
function createOvertimeApprovedNotification(employeeName, date, hours, reviewer) {
  return {
    type: "flex",
    altText: `✅ 您的加班申請已核准`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "✅ 加班核准通知",
            weight: "bold",
            size: "xl",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#FF9800",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "text",
            text: "您的加班申請已通過審核",
            size: "sm",
            color: "#666666",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "日期",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: date,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "時數",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `${hours} 小時`,
                    wrap: true,
                    color: "#FF9800",
                    size: "sm",
                    flex: 5,
                    weight: "bold"
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "審核人",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reviewer,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              }
            ]
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "text",
            text: "💪 辛苦了！",
            size: "sm",
            color: "#FF9800",
            margin: "lg",
            align: "center",
            weight: "bold"
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "primary",
            height: "sm",
            action: {
              type: "uri",
              label: "查看詳情",
              uri: "https://eric693.github.io/check_manager_plus/"
            },
            color: "#FF9800"
          }
        ],
        flex: 0
      }
    }
  };
}

/**
 * ❌ 加班審核通知（拒絕）
 */
function createOvertimeRejectedNotification(employeeName, date, hours, reviewer, reason) {
  return {
    type: "flex",
    altText: `❌ 您的加班申請已被拒絕`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: "❌ 加班未核准",
            weight: "bold",
            size: "xl",
            color: "#FFFFFF",
            align: "center"
          }
        ],
        backgroundColor: "#9E9E9E",
        paddingAll: "20px"
      },
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "text",
            text: `${employeeName}，您好！`,
            size: "lg",
            weight: "bold",
            margin: "md"
          },
          {
            type: "text",
            text: "您的加班申請未通過審核",
            size: "sm",
            color: "#666666",
            margin: "md"
          },
          {
            type: "separator",
            margin: "lg"
          },
          {
            type: "box",
            layout: "vertical",
            margin: "lg",
            spacing: "sm",
            contents: [
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "日期",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: date,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "時數",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: `${hours} 小時`,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                contents: [
                  {
                    type: "text",
                    text: "審核人",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reviewer,
                    wrap: true,
                    color: "#333333",
                    size: "sm",
                    flex: 5
                  }
                ]
              },
              {
                type: "separator",
                margin: "md"
              },
              {
                type: "box",
                layout: "baseline",
                spacing: "sm",
                margin: "md",
                contents: [
                  {
                    type: "text",
                    text: "拒絕原因",
                    color: "#999999",
                    size: "sm",
                    flex: 2
                  },
                  {
                    type: "text",
                    text: reason || "未提供",
                    wrap: true,
                    color: "#9E9E9E",
                    size: "sm",
                    flex: 5
                  }
                ]
              }
            ]
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "button",
            style: "link",
            height: "sm",
            action: {
              type: "uri",
              label: "重新申請",
              uri: "https://eric693.github.io/check_manager_plus/"
            }
          }
        ],
        flex: 0
      }
    }
  };
}

// ==================== 觸發通知函式 ====================

/**
 * 發送忘記打卡提醒
 */
function notifyForgotPunch(userId, employeeName, date, punchType) {
  const message = createForgotPunchNotification(employeeName, date, punchType);
  return sendLineNotification_(userId, message);
}

/**
 * 發送補打卡審核結果通知
 */
function notifyPunchReview(userId, employeeName, date, time, punchType, reviewer, isApproved, reason = "") {
  const message = isApproved 
    ? createPunchApprovedNotification(employeeName, date, time, punchType, reviewer)
    : createPunchRejectedNotification(employeeName, date, time, punchType, reviewer, reason);
  
  return sendLineNotification_(userId, message);
}

/**
 * 發送請假審核結果通知
 */
function notifyLeaveReview(userId, employeeName, leaveType, startDate, endDate, days, reviewer, isApproved, reason = "") {
  const message = isApproved
    ? createLeaveApprovedNotification(employeeName, leaveType, startDate, endDate, days, reviewer)
    : createLeaveRejectedNotification(employeeName, leaveType, startDate, endDate, days, reviewer, reason);
  
  return sendLineNotification_(userId, message);
}

/**
 * 發送加班審核結果通知
 */
function notifyOvertimeReview(userId, employeeName, date, hours, reviewer, isApproved, reason = "") {
  const message = isApproved
    ? createOvertimeApprovedNotification(employeeName, date, hours, reviewer)
    : createOvertimeRejectedNotification(employeeName, date, hours, reviewer, reason);
  
  return sendLineNotification_(userId, message);
}

// ==================== 定時檢查忘記打卡 ====================

/**
 * 檢查是否為平日（週一到週五）
 * @param {Date} date - 要檢查的日期
 * @returns {boolean} - true 表示是平日，false 表示是週末
 */
function isWeekday(date) {
  const day = date.getDay();
  return day >= 1 && day <= 5; // 1=週一, 5=週五
}

/**
 * 每日早上檢查昨天忘記下班打卡（只檢查平日）
 * 設定觸發器：每天早上 9:00 執行
 */
function checkForgotPunchDaily() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  // ✅ 新增：檢查昨天是否為平日
  if (!isWeekday(yesterday)) {
    Logger.log(`⏭️ ${Utilities.formatDate(yesterday, "GMT+8", "yyyy-MM-dd")} 是週末，跳過檢查`);
    return;
  }
  
  const dateStr = Utilities.formatDate(yesterday, "GMT+8", "yyyy-MM-dd");
  
  const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  
  if (!attendanceSheet || !employeeSheet) {
    Logger.log("❌ 找不到必要的工作表");
    return;
  }
  
  const employees = employeeSheet.getDataRange().getValues();
  const attendances = attendanceSheet.getDataRange().getValues();
  const headers = attendances[0];
  
  Logger.log(`📅 開始檢查 ${dateStr} (平日) 的下班打卡`);
  
  // 遍歷所有員工
  for (let i = 1; i < employees.length; i++) {
    const userId = employees[i][EMPLOYEE_COL.USER_ID];
    const name = employees[i][EMPLOYEE_COL.NAME];
    const status = employees[i][EMPLOYEE_COL.STATUS];
    
    if (status !== '啟用') continue;
    
    // 檢查昨天的打卡記錄
    let hasPunchOut = false;
    
    for (let j = 1; j < attendances.length; j++) {
      const recordDate = formatDate(attendances[j][0]);
      const recordUserId = attendances[j][1];
      const recordType = attendances[j][4]; // 打卡類別
      
      if (recordUserId === userId && recordDate === dateStr && recordType === '下班') {
        hasPunchOut = true;
        break;
      }
    }
    
    // 發送通知
    if (!hasPunchOut) {
      try {
        notifyForgotPunch(userId, name, dateStr, "下班");
        Logger.log(`📤 已提醒 ${name} 昨天忘記下班打卡`);
      } catch (err) {
        Logger.log(`⚠️ 提醒 ${name} 失敗: ${err.message}`);
      }
    }
  }
  
  Logger.log("✅ 下班打卡檢查完成");
}

/**
 * 每日早上檢查昨天忘記上班打卡（只檢查平日）
 * 設定觸發器：每天早上 9:00 執行
 */
function checkForgotPunchInMorning() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  // ✅ 新增：檢查昨天是否為平日
  if (!isWeekday(yesterday)) {
    Logger.log(`⏭️ ${Utilities.formatDate(yesterday, "GMT+8", "yyyy-MM-dd")} 是週末，跳過檢查`);
    return;
  }
  
  const dateStr = Utilities.formatDate(yesterday, "GMT+8", "yyyy-MM-dd");
  
  const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  
  if (!attendanceSheet || !employeeSheet) {
    Logger.log("❌ 找不到必要的工作表");
    return;
  }
  
  const employees = employeeSheet.getDataRange().getValues();
  const attendances = attendanceSheet.getDataRange().getValues();
  
  Logger.log(`📅 開始檢查 ${dateStr} (平日) 的上班打卡`);
  
  for (let i = 1; i < employees.length; i++) {
    const userId = employees[i][EMPLOYEE_COL.USER_ID];
    const name = employees[i][EMPLOYEE_COL.NAME];
    const status = employees[i][EMPLOYEE_COL.STATUS];
    
    if (status !== '啟用') continue;
    
    let hasPunchIn = false;
    
    for (let j = 1; j < attendances.length; j++) {
      const recordDate = formatDate(attendances[j][0]);
      const recordUserId = attendances[j][1];
      const recordType = attendances[j][4];
      
      if (recordUserId === userId && recordDate === dateStr && recordType === '上班') {
        hasPunchIn = true;
        break;
      }
    }
    
    if (!hasPunchIn) {
      try {
        notifyForgotPunch(userId, name, dateStr, "上班");
        Logger.log(`📤 已提醒 ${name} 昨天忘記上班打卡`);
      } catch (err) {
        Logger.log(`⚠️ 提醒 ${name} 失敗: ${err.message}`);
      }
    }
  }
  
  Logger.log("✅ 上班打卡檢查完成");
}

// ==================== 測試函式 ====================

/**
 * 測試忘記打卡通知
 */
function testForgotPunchNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const testDate = "2025-10-12";
  
  Logger.log("📤 測試發送忘記打卡通知...");
  const result = notifyForgotPunch(testUserId, testName, testDate, "上班");
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試補打卡審核通知（核准）
 */
function testPunchApprovedNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const testDate = "2025-10-12";
  const testTime = "09:00";
  const reviewer = "管理員";
  
  Logger.log("📤 測試發送補打卡核准通知...");
  const result = notifyPunchReview(testUserId, testName, testDate, testTime, "上班", reviewer, true);
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試補打卡審核通知（拒絕）
 */
function testPunchRejectedNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const testDate = "2025-10-12";
  const testTime = "09:00";
  const reviewer = "管理員";
  const reason = "時間不符，請重新申請";
  
  Logger.log("📤 測試發送補打卡拒絕通知...");
  const result = notifyPunchReview(testUserId, testName, testDate, testTime, "上班", reviewer, false, reason);
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試請假審核通知（核准）
 */
function testLeaveApprovedNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const leaveType = "特休假";
  const startDate = "2025-10-15";
  const endDate = "2025-10-17";
  const days = 3;
  const reviewer = "管理員";
  
  Logger.log("📤 測試發送請假核准通知...");
  const result = notifyLeaveReview(testUserId, testName, leaveType, startDate, endDate, days, reviewer, true);
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試請假審核通知（拒絕）
 */
function testLeaveRejectedNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const leaveType = "特休假";
  const startDate = "2025-10-15";
  const endDate = "2025-10-17";
  const days = 3;
  const reviewer = "管理員";
  const reason = "該時段人力不足，請調整日期";
  
  Logger.log("📤 測試發送請假拒絕通知...");
  const result = notifyLeaveReview(testUserId, testName, leaveType, startDate, endDate, days, reviewer, false, reason);
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試加班審核通知（核准）
 */
function testOvertimeApprovedNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const date = "2025-10-12";
  const hours = 3;
  const reviewer = "管理員";
  
  Logger.log("📤 測試發送加班核准通知...");
  const result = notifyOvertimeReview(testUserId, testName, date, hours, reviewer, true);
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試加班審核通知（拒絕）
 */
function testOvertimeRejectedNotification() {
  const testUserId = "U7211ffe337b29ad1f738815cb8bfdf81";
  const testName = "測試員工";
  const date = "2025-10-12";
  const hours = 3;
  const reviewer = "管理員";
  const reason = "未事先申請，請下次提前告知";
  
  Logger.log("📤 測試發送加班拒絕通知...");
  const result = notifyOvertimeReview(testUserId, testName, date, hours, reviewer, false, reason);
  Logger.log(result.ok ? "✅ 通知發送成功" : "❌ 通知發送失敗: " + result.error);
}

/**
 * 測試所有通知（一次執行所有測試）
 */
function testAllNotifications() {
  Logger.log("========== 開始測試所有通知類型 ==========\n");
  
  testForgotPunchNotification();
  Utilities.sleep(1000);
  
  testPunchApprovedNotification();
  Utilities.sleep(1000);
  
  testPunchRejectedNotification();
  Utilities.sleep(1000);
  
  testLeaveApprovedNotification();
  Utilities.sleep(1000);
  
  testLeaveRejectedNotification();
  Utilities.sleep(1000);
  
  testOvertimeApprovedNotification();
  Utilities.sleep(1000);
  
  testOvertimeRejectedNotification();
  
  Logger.log("\n========== 所有測試完成 ==========");
}