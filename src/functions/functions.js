/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
function logMessage(message) {
  console.log(message);

  return message;
}

const SQL_MI_SKUS = [
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":4, "ram":20, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":2048, "costPerHour":1.01},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":8, "ram":40, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":4096, "costPerHour":2.02},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":16, "ram":81, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":4.03},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":24, "ram":122, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":6.05},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":32, "ram":163, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":8.07},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":40, "ram":204, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":10.09},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":64, "ram":326, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":16.14},
  {"name":"General Purpose - Standard Series (Gen5)", "typeId":1, "vCore":80, "ram":408, "ramMultiplier":5.1, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":20.17},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":4, "ram":28, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":2048, "costPerHour":1.10},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":8, "ram":56, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":4096, "costPerHour":2.21},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":16, "ram":112, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":4.42},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":24, "ram":168, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":6.63},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":32, "ram":224, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":8.84},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":40, "ram":280, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":11.05},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":64, "ram":448, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":17.68},
  {"name":"General Pupose - Premium-Series", "typeId":2, "vCore":80, "ram":560, "ramMultiplier":7, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":22.08},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":4, "ram":54, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":2048, "costPerHour":1.38},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":8, "ram":108, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":4096, "costPerHour":2.77},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":16, "ram":217, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":5.54},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":24, "ram":326, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":8.31},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":32, "ram":435, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":11.08},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":40, "ram":544, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":13.85},
  {"name":"General Pupose - Memory Optimized Premium-Series", "typeId":3, "vCore":64, "ram":870, "ramMultiplier":13.6, "iopsMax":7500, "iopsDefault":500, "throughputMax":250, "storageMaxInMb":16384, "costPerHour":22.16},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":4, "ram":20, "ramMultiplier":5.1, "iopsMax":6400, "iopsDefault":300, "throughputMax":145, "storageMaxInMb":2048, "costPerHour":1.01},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":8, "ram":40, "ramMultiplier":5.1, "iopsMax":12800, "iopsDefault":300, "throughputMax":290, "storageMaxInMb":4096, "costPerHour":2.02},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":16, "ram":81, "ramMultiplier":5.1, "iopsMax":25600, "iopsDefault":300, "throughputMax":600, "storageMaxInMb":16384, "costPerHour":4.03},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":24, "ram":122, "ramMultiplier":5.1, "iopsMax":38400, "iopsDefault":300, "throughputMax":730, "storageMaxInMb":16384, "costPerHour":6.05},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":32, "ram":163, "ramMultiplier":5.1, "iopsMax":51200, "iopsDefault":300, "throughputMax":865, "storageMaxInMb":32768, "costPerHour":8.07},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":40, "ram":204, "ramMultiplier":5.1, "iopsMax":64000, "iopsDefault":300, "throughputMax":1152, "storageMaxInMb":32768, "costPerHour":10.09},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":64, "ram":326, "ramMultiplier":5.1, "iopsMax":76800, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":16.14},
  {"name":"Next-Gen General Pupose - Standard-Series (Gen5)", "typeId":4, "vCore":80, "ram":408, "ramMultiplier":5.1, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":20.17},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":4, "ram":28, "ramMultiplier":7, "iopsMax":6400, "iopsDefault":300, "throughputMax":145, "storageMaxInMb":2048, "costPerHour":1.10},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":6, "ram":42, "ramMultiplier":7, "iopsMax":9600, "iopsDefault":300, "throughputMax":217, "storageMaxInMb":2048, "costPerHour":1.65},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":8, "ram":56, "ramMultiplier":7, "iopsMax":12800, "iopsDefault":300, "throughputMax":290, "storageMaxInMb":8192, "costPerHour":2.21},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":10, "ram":70, "ramMultiplier":7, "iopsMax":16000, "iopsDefault":300, "throughputMax":367, "storageMaxInMb":8192, "costPerHour":2.76},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":12, "ram":84, "ramMultiplier":7, "iopsMax":19200, "iopsDefault":300, "throughputMax":445, "storageMaxInMb":8192, "costPerHour":3.31},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":16, "ram":112, "ramMultiplier":7, "iopsMax":25600, "iopsDefault":300, "throughputMax":600, "storageMaxInMb":16384, "costPerHour":4.42},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":20, "ram":140, "ramMultiplier":7, "iopsMax":32000, "iopsDefault":300, "throughputMax":665, "storageMaxInMb":16384, "costPerHour":5.53},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":24, "ram":168, "ramMultiplier":7, "iopsMax":38400, "iopsDefault":300, "throughputMax":730, "storageMaxInMb":16384, "costPerHour":6.64},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":32, "ram":224, "ramMultiplier":7, "iopsMax":51200, "iopsDefault":300, "throughputMax":865, "storageMaxInMb":32768, "costPerHour":8.86},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":40, "ram":280, "ramMultiplier":7, "iopsMax":64000, "iopsDefault":300, "throughputMax":1152, "storageMaxInMb":32768, "costPerHour":11.08},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":48, "ram":336, "ramMultiplier":7, "iopsMax":76800, "iopsDefault":300, "throughputMax":1152, "storageMaxInMb":32768, "costPerHour":13.30},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":56, "ram":392, "ramMultiplier":7, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":15.53},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":64, "ram":448, "ramMultiplier":7, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":17.76},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":80, "ram":560, "ramMultiplier":7, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":22.20},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":96, "ram":560, "ramMultiplier":7, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":26.64},
  {"name":"Next-Gen General Pupose - Premium-Series", "typeId":5, "vCore":128, "ram":560, "ramMultiplier":7, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":35.52},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":4, "ram":54, "ramMultiplier":13.6, "iopsMax":6400, "iopsDefault":300, "throughputMax":145, "storageMaxInMb":2048, "costPerHour":1.38},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":6, "ram":81, "ramMultiplier":13.6, "iopsMax":9600, "iopsDefault":300, "throughputMax":217, "storageMaxInMb":2048, "costPerHour":2.08},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":8, "ram":108, "ramMultiplier":13.6, "iopsMax":12800, "iopsDefault":300, "throughputMax":290, "storageMaxInMb":8192, "costPerHour":2.77},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":10, "ram":136, "ramMultiplier":13.6, "iopsMax":16000, "iopsDefault":300, "throughputMax":367, "storageMaxInMb":8192, "costPerHour":3.46},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":12, "ram":163, "ramMultiplier":13.6, "iopsMax":19200, "iopsDefault":300, "throughputMax":445, "storageMaxInMb":8192, "costPerHour":4.15},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":16, "ram":217, "ramMultiplier":13.6, "iopsMax":25600, "iopsDefault":300, "throughputMax":600, "storageMaxInMb":16384, "costPerHour":5.52},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":20, "ram":272, "ramMultiplier":13.6, "iopsMax":32000, "iopsDefault":300, "throughputMax":665, "storageMaxInMb":16384, "costPerHour":6.70},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":24, "ram":326, "ramMultiplier":13.6, "iopsMax":38400, "iopsDefault":300, "throughputMax":730, "storageMaxInMb":16384, "costPerHour":7.98},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":32, "ram":435, "ramMultiplier":13.6, "iopsMax":51200, "iopsDefault":300, "throughputMax":865, "storageMaxInMb":32768, "costPerHour":10.65},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":40, "ram":544, "ramMultiplier":13.6, "iopsMax":64000, "iopsDefault":300, "throughputMax":1152, "storageMaxInMb":32768, "costPerHour":13.32},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":48, "ram":652, "ramMultiplier":13.6, "iopsMax":76800, "iopsDefault":300, "throughputMax":1152, "storageMaxInMb":32768, "costPerHour":15.99},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":56, "ram":761, "ramMultiplier":13.6, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":18.66},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":64, "ram":870, "ramMultiplier":13.6, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":21.33},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":80, "ram":870, "ramMultiplier":13.6, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":26.66},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":96, "ram":870, "ramMultiplier":13.6, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":32.00},
  {"name":"Next-Gen General Pupose - Memory Optimized Premium-Series", "typeId":6, "vCore":128, "ram":870, "ramMultiplier":13.6, "iopsMax":80000, "iopsDefault":300, "throughputMax":1200, "storageMaxInMb":32768, "costPerHour":42.66},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":4, "ram":20, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":2.72},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":8, "ram":40, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":5.44},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":16, "ram":81, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":10.88},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":24, "ram":122, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":16.32},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":32, "ram":163, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":21.76},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":40, "ram":204, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":27.20},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":64, "ram":326, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":43.52},
  {"name":"Business Critical - Standard Series (Gen5)", "typeId":7, "vCore":80, "ram":408, "ramMultiplier":5.1, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":54.40},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":4, "ram":28, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":2.91},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":6, "ram":42, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":4.37},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":8, "ram":56, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":5.83},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":10, "ram":70, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":7.29},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":12, "ram":84, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":8.75},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":16, "ram":112, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":11.67},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":20, "ram":140, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":14.58},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":24, "ram":168, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":17.50},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":32, "ram":224, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":23.33},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":40, "ram":280, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":29.17},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":48, "ram":336, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":35.00},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":56, "ram":392, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":40.83},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":64, "ram":448, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":46.67},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":80, "ram":560, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":58.33},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":96, "ram":560, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":70.00},
  {"name":"Business Critical - Premium Series", "typeId":8, "vCore":128, "ram":560, "ramMultiplier":7, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":93.33},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":4, "ram":54, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":3.47},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":6, "ram":81, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":1024, "costPerHour":5.21},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":8, "ram":108, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":6.95},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":10, "ram":136, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":8.69},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":12, "ram":163, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":2048, "costPerHour":10.43},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":16, "ram":217, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":13.91},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":20, "ram":272, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":4096, "costPerHour":17.39},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":24, "ram":326, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":20.87},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":32, "ram":435, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":27.83},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":40, "ram":544, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":5632, "costPerHour":34.78},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":48, "ram":652, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":12288, "costPerHour":41.74},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":56, "ram":761, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":12288, "costPerHour":48.69},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":64, "ram":870, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":16384, "costPerHour":55.65},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":80, "ram":870, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":16384, "costPerHour":69.57},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":96, "ram":870, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":16384, "costPerHour":83.48},
  {"name":"Business Critical - Memory Optimized Premium-Series", "typeId":9, "vCore":128, "ram":870, "ramMultiplier":13.6, "iopsMax":120000, "iopsDefault":120000, "throughputMax":120000, "storageMaxInMb":16384, "costPerHour":111.30},
];

/**
 * Finds the cheapest SQL MI SKU that satisfies all provided criteria. Pass 0 to skip a filter.
 * @customfunction
 * @param {number} vCore Exact vCore count (0 to skip)
 * @param {number} ram Minimum RAM in GB (0 to skip)
 * @param {number} iopsMax Minimum iopsMax (0 to skip)
 * @param {number} storageInMb Minimum storageMaxInMb (0 to skip)
 * @returns {string} The name of the cheapest matching SKU, or "Match not found"
 */
function findSqlMiSkus(vCore, ram, iopsMax, storageInMb) {
  if (!vCore && !ram && !iopsMax && !storageInMb) return "";
  const matches = SQL_MI_SKUS.filter(sku =>
    (!vCore || sku.vCore >= vCore) &&
    (!ram || sku.ram >= ram) &&
    (!iopsMax || sku.iopsMax >= iopsMax) &&
    (!storageInMb || sku.storageMaxInMb >= storageInMb)
  );
  if (matches.length === 0) return "Match not found";
  const cheapest = matches.reduce((best, sku) => sku.costPerHour < best.costPerHour ? sku : best);
  return cheapest.name + " | vCore: " + cheapest.vCore + " | $" + cheapest.costPerHour + "/hr";
}

/**
 * Calculates the appropriate SQL Managed Instance configuration based on the provided parameters.
 * @customfunction
 * @param {string} CpuString Specify the CPU Name
 * @param {number} Memory Amount of Memory in MB
 * @param {number} StorageSizeInMb Specify the storage space required in MB
 * @param {number} IOPS Specify the IOPS required
 * @returns {Promise<string>} Suggest SQL Managed Instance configuration.
 */
async function calculateSqlMi(CpuString, Memory, StorageSizeInMb, IOPS) {
  try {
    const response = await fetch("https://localhost:7181/api/CalculateSqlMi", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ CpuString, Memory, StorageSizeInMb, IOPS }),
    });
    if (!response.ok) {
      return `Error: HTTP ${response.status} ${response.statusText}`;
    }
    const data = await response.json();
    return data.result;
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

