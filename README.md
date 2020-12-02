# 进来吧，这是就是你想要的百度爬虫，必能运行！可以实现百度文库自动化爬取，支持ppt,pdf,doc
## 项目介绍
#### 本项目是合法项目，只是进行数据解析而已，不能下载看不到的内容．部分文档在电脑端不能预览，但是在手机端可以预览，所有本项目把浏览器浏览格式改成手机端，支持Windows和Ubuntu．
####  本项目使用的是chromedriver来控制chrome来模拟人来操作来进行文档爬去，可以下载doc，ppt，pdf．
####  对于doc文档可以下载，doc中的表格无法下载，图片格式的文档也可以下载．
####  ppt和pdf是先下载图片再放到ppt中．
####  只要是可以预览的都可以下载．
## 已有功能.
* [x] 将可以预览的word文档下载为word文档，如果文档是扫描件，同样支持．
* [x] 将可以预览的ppt和pdf下载为不可编辑的ppt，因为网页上只有图片，所以理论上无法下载可编辑的版本．
* [ ] 支持表格下载，目前文档中的表格在网页源码中排列混乱，同时还需要结合CSS来进行布局，后续会想别的方法．
* [ ] 支持excel表格下载，目前还没有尝试，后续会试一试 ．
## 环境安装
#### pip install requests
#### pip install my_fake_useragent
#### pip install python-docx
#### pip install python-opencv
#### pip install python-pptx
#### pip install selenium
#### pip install scrapy
#### 查看自己的chrome浏览器版本，然后到chromedriver下载地址：http://npm.taobao.org/mirrors/chromedriver/ 这个地址下载对应版本的chromedriver,比如你的浏览器版本是87.0.4280，你就可以找到87.0.4280.20/这个链接，如果你是windows版本然后选择chromedriver_win32.zip进行下载解压。千万不要下载LASEST——RELEASE87.0.4280这个链接，这个链接没有用，之前有小伙伴走过弯路的，注意一下哈。
#### 本项目使用的是chromedriver控制chrome浏览器进行数据爬取的的，chromedriver的版本和chrome需要匹配，下载好的chrome放置于当前脚本同级目录，我上传的ubuntu的特定版本，windows 系统需要根据chrome版本自行下载替换。
#### windows用户下载下来的文件解压后名为chromedriver.exe, 需要把chromedriver_path = "./chromedriver_ubuntu" 改为chromedriver_path = "./chromedriver.exe",别忘了加.exe
## 使用方式：
#### 把代码中的url改为你想要下载的链接地址，脚本会自动文档判断类型，并把在当前目录新建文件夹并把文件下载到当前目录．
### 如果有需要的小伙伴需要支持可以加我QQ:2240984380,或者提isuues, 我会全力支持的！
## 如果在使用过程中遇到问题，或者有好的建议，可以在issue中提出来，我每天都会回复的，代码粗糙，大家谅解！如果好用的话就给个star
