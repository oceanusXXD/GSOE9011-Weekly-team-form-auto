# GSOE9011 Weekly Team Form 自动化填表

批量填写 GSOE9011 每周组内评价表，注：脚本仅开源作为学习用途，不提倡真实使用且作者不承担任何责任

## 一、固定填写规则

- `Week of Term` 作为输入参数
- `Group name` 作为输入参数
- 组员姓名不手填，脚本会根据你给的组号自动展开该组全部成员
- 第 4 到第 8 题固定选 `5`
- 第 9 题固定输入：

```text
Exceptional performance in all areas.
```

- 第 10 题固定选 `No`
- 实际提交时会在每次请求之间加入随机等待
- 运行日志是双语提示：中文 + English

## 二、准备环境

### 1. 安装依赖

```bash
cd /home/coder/data/GSOE9011-Weekly-team-form-auto
pip install -r requirements.txt
```

### 2. 放入浏览器 Cookie

把你从浏览器里复制出来的整串 Cookie 放到当前目录下的 `forms_cookie.txt`。

文件路径：

```text
/home/coder/data/GSOE9011-Weekly-team-form-auto/forms_cookie.txt
```

当前目录的 `.gitignore` 已经忽略了这个文件，不会误提交。

## 三、如何提取 Cookie

推荐用浏览器开发者工具，步骤如下。

### 方法 A：Network 面板

1. 用浏览器登录你的 Microsoft 账号
2. 打开表单页面
3. 按 `F12`
4. 打开 `Network`
5. 刷新页面
6. 点开 `responsepage.aspx?...` 这个请求
7. 在 `Request Headers` 里找到 `cookie`
8. 复制整行 cookie 的值，粘贴到 `forms_cookie.txt`

### 方法 B：Application / Storage 面板

1. 打开表单页面后按 `F12`
2. 进入 `Application` 或 `Storage`
3. 打开 `Cookies`
4. 选择 `https://forms.office.com`
5. 找到当前会话里的 cookie
6. 如果你不想逐个复制，还是建议回到 `Network` 里直接复制整条 `cookie` 请求头

### 提示

- Cookie 过期后，脚本会提示重新刷新 `forms_cookie.txt`
- 如果脚本突然拿不到题目结构，通常就是 Cookie 失效了

## 四、目录说明

主脚本：

```text
/home/coder/data/GSOE9011-Weekly-team-form-auto/main.py
```

依赖文件：

```text
/home/coder/data/GSOE9011-Weekly-team-form-auto/requirements.txt
```

运行过程中会自动导出这些调试文件：

- `form_page.html`
- `form_page_context.json`
- `form_payload.json`
- `form_startup.json`
- `form_schema.json`

其中 `form_schema.json` 是精简后的题目结构，适合排查题号和选项。

## 五、如何使用

### 1. 只导出表单结构，不生成提交计划

```bash
python3 /home/coder/data/GSOE9011-Weekly-team-form-auto/main.py --export-only
```

### 2. 预览批量提交计划，但不真正提交

例如组号 `60`，周次 `9`：

```bash
python3 /home/coder/data/GSOE9011-Weekly-team-form-auto/main.py --group 60 --week 9 --dry-run
```

例如组号 `60`，周次 `9,10,11`：

```bash
python3 /home/coder/data/GSOE9011-Weekly-team-form-auto/main.py --group 60 --week 9,10,11 --dry-run
```

### 3. 直接运行后手动输入

这是你最常用的模式。直接运行：

```bash
python3 /home/coder/data/GSOE9011-Weekly-team-form-auto/main.py
```

脚本会在终端依次问你：

1. `Week 几？`
2. `Group 几？`

你输入完以后，脚本就会：

1. 先定位你输入的 `Week`
2. 再定位你输入的 `Group`
3. 自动找到第三题对应的组员列表
4. 对该组所有成员自动循环
5. 后面的固定题按预设自动填写
6. 在真正提交前，先把成员列表和固定答案打印出来给你确认
7. 只有你输入 `yes`、`Y` 或 `y`，脚本才会真正开始提交

### 4. 真正执行提交

去掉 `--dry-run` 就会真实提交：

```bash
python3 /home/coder/data/GSOE9011-Weekly-team-form-auto/main.py --group 60 --week 9
```

但真正提交前，脚本还会再做一次确认：

- 展示成员列表
- 展示固定答案内容
- 等你输入 `yes`、`Y` 或 `y`

否则会直接取消，不会提交。

## 六、参数格式

`group` 支持：

- `60`
- `060`
- `Group 60`

脚本内部会自动规范成：

- `Group 01`
- `Group 60`
- `Group 154`

`week` 支持：

- `9`
- `Week 9`
- `9,10,11`

也就是你可以一次传多个 week。

## 七、提交逻辑

如果你传入：

- `group = 60`
- `week = 9,10`

脚本会：

1. 先定位 `Group 60`
2. 自动找到 `Group Member Name (Group 60)` 对应的成员列表
3. 对你输入的每个 week
4. 再对该组的每个成员
5. 生成一条提交

也就是：

```text
week 列表 × 组员列表
```

如果这个组有 6 个成员，而你传了 2 个 week，那么总共会提交 `12` 次。

## 八、注意事项

- 建议先用 `--dry-run` 看清楚提交计划和首条 payload
- 如果脚本提示 `Please refresh forms_cookie.txt`，说明当前登录态已经失效
- 随机等待时间可以自己调：

```bash
python3 /home/coder/data/GSOE9011-Weekly-team-form-auto/main.py --group 60 --week 9 --delay-min 2 --delay-max 5
```

- 这个脚本不会在自动提交，只有运行且不带 `--dry-run` 时才会真实提交

## 九、免责声明
- 这个脚本仅供学习和研究使用，请勿用于任何商业或非法用途
- 使用前请确保你已经了解并同意 Microsoft Forms 的使用条款
- 任何因使用本脚本而导致的账号问题或数据丢失，作者不承担任何责任
- 请勿滥用自动化工具，合理使用以避免对系统造成不必要的负担
- This script is intended for educational and research purposes only. Do not use it for any commercial or illegal activities.
- Please ensure you understand and agree to Microsoft Forms' terms of service before using this script.
- The author is not responsible for any account issues or data loss resulting from the use of this script.
- Do not abuse automation tools; use them responsibly to avoid unnecessary strain on the system.
