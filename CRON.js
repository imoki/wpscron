/*
    作者: imoki
    仓库: https://github.com/imoki/
    公众号：默库
    更新时间：20240719
    脚本：CRON.js 主程序，动态修改定时任务时间
    说明：再运行此CRON脚本前，请先运行CRON_INIT脚本，并配置好CRON表格的内容。
          将CRON.js加入定时任务即可自动修改定时任务时间。
*/

// 修改名称为“wps”表内的值，需要填“wps_sid”，wps_sid抓包获得

// 不要修改代码，修改wps表表格内的值即可
var filename = "" // 文件名
var cookie = ""
var file_id = file_id // 文件id
var cronArray = []  // 存放定时任务
let sheetNameSubConfig = "wps"; // 分配置表名称
let sheetNameCron = "CRON"
var pushHM = [] // 记录PUSH任务的推送时间
var hourMin = 0
var hourMax = 23
var line = 100

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}

// 获取wps_sid、cookie
function getWpsSid(){
  flagConfig = ActivateSheet(sheetNameSubConfig); // 激活wps配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("🍳 开始读取wps配置表");
    for (let i = 2; i <= 100; i++) {
      wps_sid = Application.Range("A" + i).Text; // 以第一个wps为准
      // name = Application.Range("H" + i).Text;
      break
    }
  }
  cookie = "wps_sid=" + wps_sid
  // filename = name
}

// 是否排除文件
function juiceExclude(script_name){
  let flagExclude = 0
  let i = 2
  let key = Application.Range("I" + i).Text;
  let keyarry= key.split("&") // 使用|作为分隔符
  for(let j = 0; j < keyarry.length; j ++){
    if(script_name == keyarry[j]){ // 默认排除定时任务为CRON 和PUSH的脚本
      flagExclude = 1
      console.log( "🍳 排除任务：" , keyarry[j])
      break
    }
  }
  return flagExclude
}

// 时间范围
function rangeHM(value){
  let rule = /~/i;
  let flagTrue = rule.test(value); // 判断是否存在字符串
  if (flagTrue == true) {
    console.log("🍳 使用 规则1-例如：8~13 进行时间生成")
    return 1
  } 

  rule = /\?/i;
  flagTrue = rule.test(value); // 判断是否存在字符串
  if (flagTrue == true) {
    console.log("🍳 使用 规则3-例如：6:?&?:?&?:30 进行时间生成")
    return 3
  } 

  rule = /&/i;
  flagTrue = rule.test(value); // 判断是否存在字符串
  if (flagTrue == true) {
    console.log("🍳 使用 规则2-例如：8&10&11 进行时间生成")
    return 2
  } 

  // 仅有“:”  如：8:10
  rule = /:/i;
  flagTrue = rule.test(value); // 判断是否存在字符串
  if (flagTrue == true) {
    console.log("🍳 使用 规则2-例如：8&10&11 进行时间生成")
    return 2
  } 

  console.log("🍳 使用 规则0 随机时间生成")
  return 0
  // let keyarry= value.split("~") // 使用|作为分隔符
  // // hourMin = keyarry[0]
  // hourMin = keyarry[0]
  // hourMax = keyarry[1]
  // console.log(hourMin, hourMax)
  // console.log(hourMin)
  // hourMax = keyarry[1]
  // minute = keyarry[1]
}

// 数组字符串转整形
function arraystrToint(array){
  let result = []
  for(let i=0; i<array.length; i++){
    result.push(parseInt(array[i]))
  }
  return result
}

// 数组升序排序
function arraySortUp(value){
  value.sort(function(a, b) {
    return a - b; // 升序排序
  });
  return value
}

// 数组-字典字符串转整形
function dictarraystrToint(array){
  let result = []
  for(let i=0; i<array.length; i++){
    result.push({
        "hour" : parseInt(array[i]["hour"]),
        "minute" : parseInt(array[i]["minute"]),
      })
  }
  return result
}

// 数组-字典升序排序
function dictarraySortUp(value){
  value.sort(function(a, b) {
    // console.log(a, b)
    return a["hour"] - b["hour"]; // 升序排序
  });
  return value
}

// 生成指定范围内的随机数
function getRandomInt(min, max) {
  min = Math.ceil(min);
  max = Math.floor(max);
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// 生成时间
function createTime(hour, minute, hmrange){
  console.log("⚓ 原定时时间：", hour,":",  minute)
  hour = parseInt(hour)
  minute = parseInt(minute)
  
  // if(hour > 23){
  //   hour = 0
  // }
  // if(minute > 59 || minute < 0){
  //   minute = 0
  // }

  // 根据PUSH来智能修改推送时间
  // console.log(hour, pushHM[0])
  // 规则1：8~13生成规则
  // 时分 分别加1
  let rule = 1
  rule = rangeHM(hmrange)  // 范围限定函数
  if(rule == 1){
    let keyarry= hmrange.split("~") // 使用|作为分隔符
    // hourMin = keyarry[0]
    hourMin = parseInt(keyarry[0])
    hourMax = parseInt(keyarry[1])

    hour = hour + 1
    minute = minute + 1
    if(hour > hourMax){
      hour = hourMin
      // console.log("时间大于PUSH了")
    }
    if(minute > 59 || minute < 0){
      minute = 0
    }

  }else if(rule==2){
    // 规则2：8&10&11
    // 具体时间：8:20&10:10
    let keyarry= hmrange.split("&") // 使用&作为分隔符
    let hourarry = []
    for(let k = 0; k < keyarry.length; k++){
      hourarry.push({
        "hour" : keyarry[k].split(":")[0],
        "minute" : keyarry[k].split(":")[1],
      })
    }
    // console.log(hourarry)

    // keyarry = arraySortUp(keyarry)  // 升序排序
    // keyarry = arraystrToint(keyarry)  // 转整形
    hourarry = dictarraySortUp(hourarry)  // 升序排序
    // console.log(hourarry)
    hourarry = dictarraystrToint(hourarry)  // 转整形
    // console.log(hourarry)


    // console.log(keyarry)
    // console.log(hour)
    let flagChange = 0  // 查看时间是否变化
    // for(let j=0; j < keyarry.length; j++){
    //   let hourExpect = keyarry[j]
    //   // console.log(hourExpect)
    //   if(hour < hourExpect){
    //     // 取第一个遇到比原先大的值，就变为它
    //     hour = hourExpect
    //     flagChange = 1
    //     break
    //   }
    // }

    for(let j=0; j < hourarry.length; j++){
      let hourExpect = hourarry[j]["hour"]
      let minuteExpect = hourarry[j]["minute"]
      // console.log(hourExpect)
      if(hour < hourExpect){
        // 取第一个遇到比原先大的值，就变为它
        hour = hourExpect
        // console.log(String(minuteExpect))
        if(String(minuteExpect) == "NaN"){
          // console.log("minuteExpect 为空")
        }else{
          minute = minuteExpect
        }
        
        flagChange = 1
        break
      }
    }

    // 查找最小值
    if(!flagChange){  // 如果时间没变动， 说明当前时间已经时最大了，则置为最小值

      // 无:， 即 6&8&10
      // hour = parseInt(keyarry[0]) // 则直接置为第一个值

      // 无:， 即 6&8&10
      // 有:， 即 6:10
      let hourExpect = hourarry[0]["hour"]
      let minuteExpect = hourarry[0]["minute"]
      hour = hourExpect
      // console.log(String(minuteExpect))
      if(String(minuteExpect) == "NaN"){
        // console.log("minuteExpect 为空")
      }else{
        minute = minuteExpect
      }

    }

  }else if(rule==3){
    // 规则3：8:?&7:?&?:?

    let keyarry= hmrange.split("&") // 分隔符
    let hmarray = keyarry[0]  // 查找指定的一对时分，默认为第一个

    // &分隔代表依次变成
    for(let j=0; j < keyarry.length; j++){
      // console.log(keyarry[j].split(":")[0])
      // 先找找有没有一样的，从一样的下一个开始变时间
      if(keyarry[j].split(":")[0] == hour){ // 当前时间不是和列表一样
        flagFind = 1
        // 找到一样的了
        // if(j == keyarry.length - 1){  // 是最后一个，那么就取第一个
        //   hmarray = keyarry[0]
        //   break
        // }
        hmarray = keyarry[(j + 1) % keyarry.length] // 是最后一个，那么就取第一个
        break
      }
    }

    // 开始变值

    let array2 = hmarray.split(":")
    // console.log(array2)
    hourRandom = array2[0]
    minuteRandom = array2[1]
    // console.log(hourRandom, minuteRandom)
    if(hourRandom == "?"){
      // 随机生成时间
      hour = getRandomInt(0, 23);
      // console.log("随机生成时")
    }else{
      hour = hourRandom
    }

    if(minuteRandom == "\?"){
      // 随机生成时间
      // console.log("随机生成分")
      minute = getRandomInt(0, 60);
      // console.log(minute)
      
    }else{
      minute = minuteRandom
    }
  }else{
    // 所有规则都不是
    // 则随机生成
    hour = getRandomInt(0, 23);
    minute = getRandomInt(0, 60);
  }

    
  newHM = [hour.toString(), minute.toString()]
  return newHM
}


// 获取定时任务
function getTask(url, headers){
  // 查看定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  // list -> 数组 -> file_id、task_id、script_name，cron_detail->字典
  cronlist = resp["list"]
  // console.log(cronlist)
  for(let i = 0; i < cronlist.length; i++){
    
    task = cronlist[i]
    task_id = task["task_id"]
    script_id = task["script_id"]
    script_name = task["script_name"]

    cron_detail = task["cron_detail"]
    cron_desc = cron_detail["cron_desc"]
    cron_type = cron_desc["cron_type"]
    // day_of_month = cron_desc["day_of_month"]
    // day_of_week = cron_desc["day_of_week"]
    // month = cron_desc["month"]
    hour = cron_desc["hour"]
    minute = cron_desc["minute"]
    // year = cron_desc["year"]

    let flagExclude = 0;  // 排除掉的文件，1为排除
    flagExclude = juiceExclude(script_name)

    if(!flagExclude){   // 不排除的任务就进行修改
      console.log("✨ 原定时任务：" , script_name, " 定时时间：", hour,":",  minute)
      cronArray.push(task) // 加入任务列表
    }

    if(script_name == "PUSH"){  // 记录PUSH的推送脚本的时间
      // console.log("记录PUSH推送时间")
      pushHM[0] = parseInt(hour)
      pushHM[1] = parseInt(minute)
      // console.log(pushHM[0], pushHM[1])
    }
  }
  sleep(5000)
}

// 写入最新的时间
function writeNewTime(pos, hour, minute){
  Application.Range("H" + pos).Value = hour + ":" + minute
}


// 修改定时任务
function putTask(url, headers, data, task_id, script_name){
  let flagResult = 0
  // console.log(url + "/" + task_id)
  // 查看定时任务
  resp = HTTP.put(
    url + "/" + task_id,
    data,
    { headers: headers }
  );
  resp = resp.json()
  // console.log(resp)

  // {"result":"ok"}
  // {"errno":10000,"msg":"value of hour is out of 24's bounds","reason":"","result":"InvalidArgument"}
  result = resp["result"]
  if(result == "ok"){
    console.log("🎉 " + script_name + " 任务时间调整成功")
    flagResult = 1
  }else{
    msg = resp["msg"]
    console.log("📢 " , msg)
  }
  sleep(5000)
  return flagResult
}


function main(){
  
  getWpsSid() // 获取cookie
  headers= {
    "Cookie": cookie,
    "Content-Type" : "application/json",
    "Origin":"https://www.kdocs.cn",
    "Priority":"u=1, i",
  //   "Content-Type":"application/x-www-form-urlencoded",
  }
  // console.log(headers)

  
  // 设置定时任务
  ActivateSheet(sheetNameCron);

  let file_name = ""
  let file_id = ""
  let task_name = ""
  let script_name = ""
  let task_id  = ""
  let script_id = ""
  let exclude = ""
  let hmrange = ""
  let hm = ""
  let hour = 0
  let minute = 0
  let extra = ""
  let cron_type = ""
  let day_of_month = ""
  let day_of_week = ""
  for (let i = 2; i <= line; i++) {
      file_name = Application.Range("A" + i).Text;
      if (file_name == "") {
          // 如果为空行，则提前结束读取
          break;
      }

      exclude = Application.Range("F" + i).Text;  // 是否调整

      if (exclude == "是") {  // 是代表进行调整，则进行修改
        file_id = Application.Range("B" + i).Value;
        // console.log(file_id)
        task_name = Application.Range("C" + i).Text;
        script_name = task_name
        task_id = Application.Range("D" + i).Text;
        script_id = Application.Range("E" + i).Text;
        hmrange = Application.Range("G" + i).Text;
        hm = Application.Range("H" + i).Text;
        extra = Application.Range("I" + i).Text;
        console.log("🧑 开始任务修改：" , file_name, "-", task_name )
        let keyarry= hm.split(":") // 使用:作为分隔符

        hour = parseInt(keyarry[0])
        minute = parseInt(keyarry[1])
        // rangeHM(hmrange)  // 范围限定函数
        newHM = createTime(hour, minute, hmrange)
        hour = newHM[0]
        minute = newHM[1]

        keyarry = extra.split("&") // 使用|作为分隔符
        cron_type = keyarry[0]
        day_of_month = keyarry[1]
        day_of_week = keyarry[2]

        // 进行时间修改
        url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/cron_tasks";
        // console.log(url)
        data = {
          "id": file_id,
          "script_id": script_id,
          "cron_detail": {
              "task_type": "cron_task",
              "cron_desc": {
                  "cron_type": cron_type,
                  "day_of_month": day_of_month,
                  "day_of_week": day_of_week,
                  "hour" : hour,
                  "minute": minute
              }
          },
          "task_id": task_id,
          "status": "enable"
        }
        // console.log(data)

        console.log("✨ 现定时任务：" , script_name, " 定时时间：", hour,":",  minute)
        let flagResult = putTask(url, headers, data, task_id, script_name)
        if(flagResult){ // 时间修改成功
          let pos = i
          writeNewTime(pos, hour, minute)
        }

      }
  } 

}

main()