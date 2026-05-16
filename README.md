# HS-Chinese-Context-Based-Dictation-Exam-Q-Generator
这是一个基于Python的高中语文情境式默写题目自动生成工具，专门为语文教师和教育教学工作者设计。该工具能够自动分析古文篇章，生成符合高考要求的情境式默写题目，并输出格式规范的Word文档。

一、核心功能

1.自动篇章分析：自动识别TXT文件中的部编版高中语文必备篇目

2.智能题目生成：利用DeepSeek API生成高质量的情境式默写题目

3.格式规范输出：自动生成包含题目和参考答案的Word文档

4.智能句子分割：严格按照标点符号（逗号、句号等）分割诗句

5.自动文件管理：自动创建输出文件夹并打开查看

二、运行环境要求

1.Python版本：3.7 或更高版本

2.操作系统：Windows 10/11, macOS, Linux

3.网络连接：需要能够访问DeepSeek API

三、打包为可执行文件

1.安装依赖

pip install -r requirements.txt

2.可以使用PyInstaller将程序打包为独立的可执行文件：

pyinstaller --onefile --name="情境默写生成器" 情境默写题目生成器.py

四、打包后使用说明

1.将生成的情境默写生成器.exe和config.json放在同一目录

2.在项目根目录创建config.json文件，内容如下：

json
{
  "deepseek_api_key": "your_api_key_here",
  "knowledge_base_path": "knowledge_base.json",
  "default_output_dir": "./output"
}

将your_api_key_here替换为您的DeepSeek API密钥。

3.双击运行情境默写生成器.exe

五、注意事项

1.API使用

（1）需要有效的DeepSeek API密钥

（2）注意API调用频率限制

（3）确保网络连接稳定

2.输入文件

（1）确保TXT文件编码为UTF-8

（2）文件中应包含完整的古文原文

（3）建议每篇古文独立成段

3.输出结果

（1）输出文件夹会自动创建

（2）每次运行生成独立的Word文档

（3）程序会自动打开输出文件夹

六、常见问题

1. API密钥错误

（1）检查config.json中的API密钥是否正确

（2）确保API密钥有足够的额度

（3）确认网络连接正常

2. 文件读取失败

（1）检查TXT文件编码是否为UTF-8

（2）确保文件路径不包含特殊字符

（3）确认文件未被其他程序占用

3. 生成题目格式错误

（1）检查输入文本格式是否规范

（2）确认诗句用正确的标点符号分隔

（3）如遇问题，可调整提示词重新生成

七、提示：本工具为教育教学辅助工具，生成的题目仅供参考，建议教师在使用前进行审查和调整。
