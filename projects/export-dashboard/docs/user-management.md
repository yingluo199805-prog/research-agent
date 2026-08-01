# 用户管理工具

这个工具用于在任意电脑上新增、查询、删除出海数据看板的内测用户。用户数据存放在 Upstash Redis：

`user:<email>` -> `{"password":"...","name":"...","org":"...","createdAt":"..."}`

## 1. 首次配置

进入仓库目录：

```powershell
cd projects\export-dashboard\tools
Copy-Item .env.example .env
notepad .env
```

把 `.env` 里的 `UPSTASH_REDIS_REST_TOKEN` 改成真实 token。

注意：`.env` 已在 `.gitignore` 中忽略，不要提交到 GitHub，也不要发给客户。

## 2. 新增用户

```powershell
.\user-manager.ps1 add -Email "someone@example.com" -Name "张三" -Org "某某基金"
```

脚本会自动生成 6 位临时密码，并输出可直接转发的邀请语。

如需指定密码：

```powershell
.\user-manager.ps1 add -Email "someone@example.com" -Name "张三" -Org "某某基金" -Password "AbC123"
```

## 3. 查询用户

```powershell
.\user-manager.ps1 get -Email "someone@example.com"
```

## 4. 重新输出邀请语

```powershell
.\user-manager.ps1 invite -Email "someone@example.com"
```

## 5. 删除用户

```powershell
.\user-manager.ps1 delete -Email "someone@example.com"
```

## 6. 列出全部用户

```powershell
.\user-manager.ps1 list
```

## 7. 另一台电脑使用

1. 从 GitHub 私有仓库 clone 或下载这个仓库。
2. 进入 `projects\export-dashboard\tools`。
3. 复制 `.env.example` 为 `.env`。
4. 填入真实 `UPSTASH_REDIS_REST_TOKEN`。
5. 运行 `user-manager.ps1`。

如果 PowerShell 提示脚本执行策略限制，可以只对当前窗口放开：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
