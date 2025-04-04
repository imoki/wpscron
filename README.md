# 🎉 金山文档智能调整定时任务时间程序
<div align="center">
    <img src="https://socialify.git.ci/imoki/wpscron/image?description=1&font=Rokkitt&forks=1&issues=1&language=1&owner=1&pattern=Circuit%20Board&pulls=1&stargazers=1&theme=Dark">
<h1>默调时</h1>
<!-- <h1>（已失效）</h1> -->
基于「金山文档」的智能调时程序

<div id="shield">

[![][github-stars-shield]][github-stars-link]
[![][github-forks-shield]][github-forks-link]
[![][github-issues-shield]][github-issues-link]
[![][github-contributors-shield]][github-contributors-link]

<!-- SHIELD GROUP -->
</div>
</div>

## 🎊 简介
此脚本能够自动对金山文档定时任务中的时间进行自动调整，以此达到定时任务每天执行时间都不一样的效果  
将CRON脚本加入定时任务，每天指定时间将会进行时间调整  
脚本能设置只调整哪些定时任务，能统一对多个文档内的多个定时任务进行个性化修改  
脚本具备多种时间模式，多种灵活功能  

## ✨ 特性
    - 📀 支持金山文档运行
    - 💿 支持普通表格和智能表格
    - ♾️ 支持限定智能调整后的时间范围
    - 💽 支持多文档统一修改
    

## 🛰️ 文字步骤
1. 将CRON_INIT、CRON脚本添加到金山文档中
2. 给CRON_INIT、CRON脚本添加网络API
3. 第一次运行CRON_INIT脚本
4. 填写自动生成的wps表中的wps_sid
5. 再次运行CRON_INIT脚本
6. 填写自动生成的CRON表中的内容
7. 将CRON脚本加入定时任务

## ⭐ 表格参考例子
![wps表](https://s3.bmp.ovh/imgs/2024/07/14/9045db168c0875ee.png)
![CRON表](https://s3.bmp.ovh/imgs/2024/07/14/dc9fcfdf5ba3eb7c.png)
![时间](https://s3.bmp.ovh/imgs/2024/07/15/5b3b7904259cc1ac.png)

## 🧾 表格内容含义 
1. wps_sid ： 填写wps文档内抓包得到的wps_sid
2. 文档名 : 填写需要修改定时任务时间的文档名称
3. 是否调整 ： 选项填“是”则会对其进行时间调整，默认为“否”是排除这个任务不会进行调整
4. 排除文档 ： 代表哪些文档不读取。以&分隔文档名，如：文档1&文档2
5. 仅读取文档 ： 代表仅读取哪些文档。以&分隔文档名，如：文档1&文档2。默认为@all代表所有文档都读取

## 🕙 时间范围填写格式
1. 规则0：随机生成时和分。什么都不填。  
1. 规则1：允许的最小时到最大时。例如:8\~13，代表调整后时间超过14点(14:xx)后会自动调整为8:xx。  
2. 规则2：以&分隔整点。如8&10&11，则会依次调整为8:xx、10:xx、11:xx点。也可指定准确时间：如8:10。则会置为8:10  
3. 规则3：以?填充，?代表随机的值。如6:?&?:?&?:30，则会依次调整为6:xx、xx:xx、xx:30点。  


## 🚀 其他
如果手动修改了定时任务时间，请重新运行一次CRON_INIT脚本，会自动生成最新的CRON配置表

## 🤝 欢迎参与贡献
欢迎各种形式的贡献

[![][pr-welcome-shield]][pr-welcome-link]

<!-- ### 💗 感谢我们的贡献者
[![][github-contrib-shield]][github-contrib-link] -->


## ✨ Star 数

[![][starchart-shield]][starchart-link]

## 📝 更新日志 
- 2024-07-16
    * 新增时间生成规则，指定准确时间
- 2024-07-15
    * 新增时间生成规则，以?填充
- 2024-07-14
    * 新增排除文档功能
    * 新增仅读取文档功能
- 2024-07-13
    * 新增时间生成规则，以&分隔整点
- 2024-07-12
    * 动态判断表格是否填写
    * 支持多文档统一修改
    * 支持时间范围限定
    * 支持普通表格和智能表格
- 2024-07-11
    * 推出金山文档智能调定时程序

## 📌 特别声明

- 本仓库发布的脚本仅用于测试和学习研究，禁止用于商业用途，不能保证其合法性，准确性，完整性和有效性，请根据情况自行判断。

- 本人对任何脚本问题概不负责，包括但不限于由任何脚本错误导致的任何损失或损害。

- 间接使用脚本的任何用户，包括但不限于建立VPS或在某些行为违反国家/地区法律或相关法规的情况下进行传播, 本人对于由此引起的任何隐私泄漏或其他后果概不负责。

- 请勿将本仓库的任何内容用于商业或非法目的，否则后果自负。

- 如果任何单位或个人认为该项目的脚本可能涉嫌侵犯其权利，则应及时通知并提供身份证明，所有权证明，我们将在收到认证文件后删除相关脚本。

- 任何以任何方式查看此项目的人或直接或间接使用该项目的任何脚本的使用者都应仔细阅读此声明。本人保留随时更改或补充此免责声明的权利。一旦使用并复制了任何相关脚本或Script项目的规则，则视为您已接受此免责声明。

**您必须在下载后的24小时内从计算机或手机中完全删除以上内容**

> ***您使用或者复制了本仓库且本人制作的任何脚本，则视为 `已接受` 此声明，请仔细阅读***

<!-- LINK GROUP -->
[github-codespace-link]: https://codespaces.new/imoki/wpscron
[github-codespace-shield]: https://github.com/imoki/wpscron/blob/main/images/codespaces.png?raw=true
[github-contributors-link]: https://github.com/imoki/wpscron/graphs/contributors
[github-contributors-shield]: https://img.shields.io/github/contributors/imoki/wpscron?color=c4f042&labelColor=black&style=flat-square
[github-forks-link]: https://github.com/imoki/wpscron/network/members
[github-forks-shield]: https://img.shields.io/github/forks/imoki/wpscron?color=8ae8ff&labelColor=black&style=flat-square
[github-issues-link]: https://github.com/imoki/wpscron/issues
[github-issues-shield]: https://img.shields.io/github/issues/imoki/wpscron?color=ff80eb&labelColor=black&style=flat-square
[github-stars-link]: https://github.com/imoki/wpscron/stargazers
[github-stars-shield]: https://img.shields.io/github/stars/imoki/wpscron?color=ffcb47&labelColor=black&style=flat-square
[github-releases-link]: https://github.com/imoki/wpscron/releases
[github-releases-shield]: https://img.shields.io/github/v/release/imoki/wpscron?labelColor=black&style=flat-square
[github-release-date-link]: https://github.com/imoki/wpscron/releases
[github-release-date-shield]: https://img.shields.io/github/release-date/imoki/wpscron?labelColor=black&style=flat-square
[pr-welcome-link]: https://github.com/imoki/wpscron/pulls
[pr-welcome-shield]: https://img.shields.io/badge/🤯_pr_welcome-%E2%86%92-ffcb47?labelColor=black&style=for-the-badge
[github-contrib-link]: https://github.com/imoki/wpscron/graphs/contributors
[github-contrib-shield]: https://contrib.rocks/image?repo=imoki%2Fsign_script
[docker-pull-shield]: https://img.shields.io/docker/pulls/imoki/wpscron?labelColor=black&style=flat-square
[docker-pull-link]: https://hub.docker.com/repository/docker/imoki/wpscron
[docker-size-shield]: https://img.shields.io/docker/image-size/imoki/wpscron?labelColor=black&style=flat-square
[docker-size-link]: https://hub.docker.com/repository/docker/imoki/wpscron
[docker-stars-shield]: https://img.shields.io/docker/stars/imoki/wpscron?labelColor=black&style=flat-square
[docker-stars-link]: https://hub.docker.com/repository/docker/imoki/wpscron
[starchart-shield]: https://api.star-history.com/svg?repos=imoki/wpscron&type=Date
[starchart-link]: https://api.star-history.com/svg?repos=imoki/wpscron&type=Date

