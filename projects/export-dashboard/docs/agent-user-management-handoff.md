# 给另一台电脑 Agent 的用户管理操作说明

目标：帮助用户管理出海数据跟踪网站的内测账号。账号数据存放在 Upstash Redis 中，通过本仓库里的 PowerShell 脚本操作。

## 重要原则

1. 不要把真实 `UPSTASH_REDIS_REST_TOKEN` 提交到 GitHub。
2. 真实 token 只应放在本机本地文件：`projects/export-dashboard/tools/.env`。
3. 如果 `.env` 不存在，或 token 还是占位符，必须向用户索要 token。
4. 用户只需要提供客户的邮箱、姓名、机构；脚本会自动生成 6 位临时密码并输出邀请语。

## 仓库位置

GitHub 私有仓库：

```text
https://github.com/yingluo199805-prog/research-agent
```

用户管理脚本位置：

```text
projects/export-dashboard/tools/user-manager.ps1
```

配置模板位置：

```text
projects/export-dashboard/tools/.env.example
```

本机真实配置文件位置：

```text
projects/export-dashboard/tools/.env
```

`.env` 已被 `.gitignore` 忽略，正常不会被提交。

## 首次在另一台电脑配置

如果仓库还没有下载：

```powershell
git clone https://github.com/yingluo199805-prog/research-agent.git
cd research-agent\projects\export-dashboard\tools
```

如果仓库已经存在：

```powershell
cd research-agent
git pull
cd projects\export-dashboard\tools
```

检查 `.env`：

```powershell
Test-Path .env
```

如果返回 `False`，创建它：

```powershell
Copy-Item .env.example .env
notepad .env
```

`.env` 应包含两行：

```env
UPSTASH_REDIS_REST_URL=https://busy-jackass-97739.upstash.io
UPSTASH_REDIS_REST_TOKEN=<真实 token>
```

如果不知道真实 token，向用户索要。可以这样问：

```text
我需要 Upstash Redis REST token 才能在这台电脑上新增/删除用户。请把 token 发给我；我只会保存到本机 tools/.env，不会提交到 GitHub。
```

## PowerShell 执行策略

如果运行脚本时报执行策略错误，先在当前 PowerShell 窗口执行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## 新增用户

用户通常会给：

```text
机构 姓名 邮箱
```

或：

```text
邮箱 姓名 机构
```

请识别出三项后执行：

```powershell
.\user-manager.ps1 add -Email "客户邮箱" -Name "客户姓名" -Org "客户机构"
```

示例：

```powershell
.\user-manager.ps1 add -Email "someone@example.com" -Name "张三" -Org "某某基金"
```

脚本会：

1. 检查该邮箱是否已存在。
2. 如果不存在，自动生成 6 位临时密码。
3. 写入 Redis。
4. 输出邀请语。

把脚本输出的邀请语原样反馈给用户。

## 查询用户

```powershell
.\user-manager.ps1 get -Email "someone@example.com"
```

## 重新输出邀请语

```powershell
.\user-manager.ps1 invite -Email "someone@example.com"
```

## 删除用户

删除前先向用户确认，因为这是破坏性操作：

```powershell
.\user-manager.ps1 delete -Email "someone@example.com"
```

## 列出所有用户

```powershell
.\user-manager.ps1 list
```

## 常见问题

### 1. Upstash 连接失败

如果报错类似：

```text
Unable to connect to the remote server
```

通常是当前网络访问 Upstash 不稳定。建议用户打开代理/梯子，必要时使用全局模式，然后重试。

### 2. `.env` 不要提交

任何时候都不要执行类似：

```powershell
git add .env
```

提交前可检查：

```powershell
git status --short --ignored projects/export-dashboard/tools/.env
```

如果看到：

```text
!! projects/export-dashboard/tools/.env
```

说明 `.env` 被忽略，是正确状态。

### 3. 邮箱大小写

脚本会把邮箱统一转成小写作为 Redis key。登录时建议也让客户直接复制邀请语中的账号。

## 邀请语模板

脚本会自动输出类似以下内容：

```text
X总好，

感谢您参加我们出海数据跟踪网站的内测，如国内访问遇到问题请连接梯子访问。

内测网址：https://dashboard.gfautodata.cn
登录账号：客户邮箱
登录密码：临时密码

出海数据网站包括中汽协乘用车出口、海关出口（分区域、分国家、分车企、分车型；独家数据）、Marklines终端实销情况等，更多功能陆续上线，数据月度更新。

欢迎多多提问题，我将进一步完善网站开发，将内测反馈意见修改后正式推出跟踪网站。

备注：请勿传播网站内数据，最终解释权归广发汽车。
```
