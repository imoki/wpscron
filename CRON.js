// CRON 动态修改定时任务时间
// 20240711

// 修改名称为“wps”表内的值，需要填“wps_sid”和“文档名”
// wps_sid抓包获得，文档名就是你这个文档的名称

// 不要修改代码，修改wps表表格内的值即可
var filename = "" // 文件名
var cookie = ""
var file_id = file_id // 文件id
var cronArray = []  // 存放定时任务
let sheetNameSubConfig = "wps"; // 分配置表名称
var pushHM = [] // 记录PUSH任务的推送时间
var hourMin = 0

// // 定时任务类型设置为每日
// cron_type = "daily"

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
      name = Application.Range("B" + i).Text;
      break
    }
  }
  cookie = "wps_sid=" + wps_sid
  filename = name
}

// 是否排除文件
function juiceExclude(script_name){
  let flagExclude = 0
  let i = 2
  let key = Application.Range("C" + i).Text;
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
function rangeHM(){
  let i = 2
  let key = Application.Range("D" + i).Text;
  let keyarry= key.split("~") // 使用|作为分隔符
  hourMin = keyarry[0]
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
  if(hour > pushHM[0]){
    hour = hourMin
    // console.log("时间大于PUSH了")
  }
  if(minute > 59 || minute < 0){
    minute = 0
  }
    
  newHM = [hour.toString(), minute.toString()]
  return newHM
}

// 获取file_id
function getFileId(url, headers, filename){
  // 查看定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  resplist = resp["list"]
  for(let i = 0; i<resplist.length; i++){
    roaming = resplist[i]["roaming"]
    // console.log(roaming)
    fileid = roaming["fileid"]
    name = roaming["name"]
    if(filename + ".xlsx" == name){
      console.log("🍳 已找到指定文档，对指定文档进行操作")
      file_id = fileid
      // console.log("🍳 file_id : " + file_id)
      break
    }
  }
  sleep(5000)
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


// 修改定时任务
function putTask(url, headers, data, task_id, script_name){
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
  }else{
    msg = resp["msg"]
    console.log("📢 " , msg)
  }
  sleep(5000)
}

// data = {
//   "id": file_id,
//   "script_id": script_id,
//   "cron_detail": {
//       "task_type": "cron_task",
//       "cron_desc": {
//           "cron_type": cron_type,
//           "day_of_month": day_of_month,
//           "day_of_week": day_of_week,
//           "hour": hour,
//           "minute": minute
//       }
//   }
// }

rangeHM() // 时间访问，超过PUSH后进行调整的时间
getWpsSid() // 获取cookie
headers= {
  "Cookie": cookie,
  "Content-Type" : "application/json",
  "Origin":"https://www.kdocs.cn",
  "Priority":"u=1, i",
//   "Content-Type":"application/x-www-form-urlencoded",
}
// console.log(headers)


// 获取file_id
url = "https://drive.kdocs.cn/api/v5/roaming?count=1"
getFileId(url, headers, filename)

// 设置定时任务
url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/cron_tasks";
// console.log(url)
getTask(url, headers)


for(let i = 0; i < cronArray.length; i++){
  task = cronArray[i]
  task_id = task["task_id"]
  script_id = task["script_id"]
  script_name = task["script_name"]

  cron_detail = task["cron_detail"]
  cron_desc = cron_detail["cron_desc"]
  cron_type = cron_desc["cron_type"]
  day_of_month = cron_desc["day_of_month"]
  day_of_week = cron_desc["day_of_week"]
  // month = cron_desc["month"]
  hour = cron_desc["hour"]
  minute = cron_desc["minute"]
  // year = cron_desc["year"]
  newHM = createTime(hour, minute)
  hour = newHM[0]
  minute = newHM[1]

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
  putTask(url, headers, data, task_id, script_name)
  // break
  
}

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}



// // {"errno":20024,"msg":"","reason":"","result":"SessionDeleted"}
// // {"task_id":"xxx"}
// resp = resp.json()
// console.log(resp)
// task_id = resp["task_id"] // 获取task_id

// // 删除定时设置任务
// // url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/cron_tasks/" + task_id
// DELETE


