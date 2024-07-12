/*
    作者: imoki
    仓库: https://github.com/imoki/
    公众号：默库
    更新时间：20240712
    脚本：CRON.js 主程序，动态修改定时任务时间
    说明：再运行此CRON脚本前，请先运行CRON_INIT脚本，并配置好CRON表格的内容。
          将CRON.js加入定时任务即可自动修改定时任务时间。
          修改规则为每次往后推1小时
*/

// 修改名称为“wps”表内的值，需要填“wps_sid”和“文档名”
// wps_sid抓包获得，文档名就是你这个文档的名称

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
  let keyarry= value.split("~") // 使用|作为分隔符
  // hourMin = keyarry[0]
  hourMin = keyarry[0]
  hourMax = keyarry[1]
  // console.log(hourMin, hourMax)
  // console.log(hourMin)
  // hourMax = keyarry[1]
  // minute = keyarry[1]
}

// 生成时间
function createTime(hour, minute){
  hour = parseInt(hour)
  minute = parseInt(minute)
  hour = hour + 1
  minute = minute + 1
  // if(hour > 23){
  //   hour = 0
  // }
  // if(minute > 59 || minute < 0){
  //   minute = 0
  // }

  // 根据PUSH来智能修改推送时间
  // console.log(hour, pushHM[0])
  if(hour > hourMax){
    hour = hourMin
    // console.log("时间大于PUSH了")
  }
  if(minute > 59 || minute < 0){
    minute = 0
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
        let keyarry= hm.split(":") // 使用|作为分隔符

        hour = parseInt(keyarry[0])
        minute = parseInt(keyarry[1])
        rangeHM(hmrange)  // 范围限定函数
        newHM = createTime(hour, minute)
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



