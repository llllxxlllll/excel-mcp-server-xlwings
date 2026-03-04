# 从零上传项目到 GitHub 指南

## 一、先搞懂几件事

- **Git**：在你电脑上用的版本管理工具（你已经装好了）。
- **GitHub**：网站，用来存代码、和别人协作。你要做的就是把本地项目“推”到 GitHub 上的一个**仓库**里。
- **仓库（Repository）**：就是 GitHub 上的一个项目文件夹，一个项目对应一个仓库。

**流程概括**：在 GitHub 上新建一个空仓库 → 在本地把代码“提交”并“推”到这个仓库。

---

## 二、第一步：在 GitHub 上建仓库

1. 打开浏览器，访问：**https://github.com**
2. 登录你的账号（没有就先注册）。
3. 点击右上角 **“+”** → **“New repository”**（新建仓库）。
4. 填写：
   - **Repository name**：仓库名，例如 `excel-mcp-server-xlwings`（只能用英文、数字、短横线）。
   - **Description**：选填，例如“基于 xlwings 的 Excel MCP 服务端”。
   - 选 **Public**（公开）。
   - **不要**勾选 “Add a README file”等（我们要推送已有代码，保持空仓库）。
5. 点 **“Create repository”**。
6. 创建好后，页面上会有一个地址，类似：
   ```text
   https://github.com/你的用户名/excel-mcp-server-xlwings.git
   ```
   复制这个地址，后面要用。

---

## 三、第二步：在本地用 Git 推送代码

打开 **PowerShell**，在项目目录下按顺序执行下面命令。

### 1. 进入项目目录（如已在项目目录可跳过）

```powershell
Set-Location "D:\res_program\excel-mcp-server-xlwings"
```

### 2. 把当前目录设为 Git 安全目录（若之前没做过）

若执行后面命令时提示 “dubious ownership”，先执行：

```powershell
git config --global --add safe.directory D:/res_program/excel-mcp-server-xlwings
```

### 3. 换掉远程仓库地址（重要）

你现在本地指向的是原作者的仓库。要上传到**你自己的**仓库，需要改成你的地址：

```powershell
git remote set-url origin https://github.com/你的用户名/你的仓库名.git
```

例如仓库名是 `excel-mcp-server-xlwings`、用户名为 `zhangsan`，则：

```powershell
git remote set-url origin https://github.com/zhangsan/excel-mcp-server-xlwings.git
```

（把 `zhangsan` 和 `excel-mcp-server-xlwings` 换成你自己的。）

### 4. 添加要提交的文件

```powershell
git add .
```

（`.` 表示当前目录下所有改动和新文件，按 `.gitignore` 规则会忽略 .venv、.hypothesis 等。）

### 5. 做一次“提交”（保存当前版本）

```powershell
git commit -m "首次上传：xlwings 版 Excel MCP 服务端"
```

（`-m` 后面是这次提交的说明，可以随便改成你喜欢的话。）

### 6. 推送到 GitHub

```powershell
git push -u origin main
```

- 若你的默认分支叫 `master` 而不是 `main`，把上面命令里的 `main` 改成 `master`。
- 第一次推送可能会弹出浏览器或窗口让你登录 GitHub 或输入账号密码；按提示完成即可。

执行完后，刷新你在 GitHub 上刚建的那个仓库页面，就能看到代码已经在了。

---

## 四、以后改了代码，再上传怎么弄？

在项目目录下执行：

```powershell
git add .
git commit -m "这里写你做了啥修改"
git push
```

---

## 五、常见问题

| 情况 | 处理 |
|------|------|
| 提示 “dubious ownership” | 执行第二步里的 `git config --global --add safe.directory ...`。 |
| 推送时要登录 | 按提示在浏览器登录 GitHub，或使用 GitHub 的 Personal Access Token 当密码。 |
| 提示 “remote origin already exists” | 用 `git remote set-url origin 你的新地址` 覆盖即可。 |
| 想看当前远程地址 | `git remote -v`。 |

---

**总结**：在 GitHub 建空仓库 → 本地改 `origin` 为你的仓库地址 → `git add .` → `git commit -m "说明"` → `git push -u origin main`。做完这遍，你就完成第一次上传了。
