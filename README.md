# Excel 表格 → 按产品编号建目录并复制客户文件

从 **Excel** 读取「产品编号」「客户型号」两列，支持手动纠错后，按产品编号创建文件夹及 YG 子目录，并从客户文件目录中按客户编号复制对应**压缩文件**（.zip/.rar/.7z 等）与 **PDF** 到各 YG。

## 环境与依赖

- Python 3.8+
- **openpyxl**（读 Excel .xlsx）

### 安装

**用 Conda（推荐）**：

```bash
conda env create -f environment.yml
conda activate file-mkdir
```

环境已包含 openpyxl。若环境已存在但缺依赖，可单独安装：

```bash
conda activate file-mkdir
conda install openpyxl
```

**用 pip**：

```bash
pip install -r requirements.txt
```

## 运行

**用 Conda 时**（先激活环境再运行）：

```bash
conda activate file-mkdir
python main.py
```

或已用 pip 装好依赖后直接：

```bash
python main.py
```

1. **选择 Excel 表格**：第一行为表头，需包含「产品编号」「客户型号」两列（.xlsx / .xls）。
2. 程序读取两列并展示，**双击单元格可修改**。
3. **选择生成目录位置**：如 `D:/测试/`，将在此下创建 `产品编号/YG`。
4. **选择客户文件所在目录**：该目录及子目录中**名称包含「客户编号」的压缩文件**（.zip、.rar、.7z、.tar.gz 等）和 **PDF** 会被复制到对应 `产品编号/YG`。
5. 点击 **开始执行**，查看日志。

命令行测试（无 GUI，需先激活环境）：

```bash
conda activate file-mkdir
python main.py --cli
```

## 打包成 exe（Windows 免安装运行）

**必须在 Windows 上打包**，得到的 exe 可拷贝到任意 Windows 电脑直接双击运行，无需安装 Python、Conda 或任何依赖。

### 在 Windows 上操作步骤

1. **把整个项目拷到 Windows**（U 盘、网盘、Git 等）。

2. **安装 Python 和依赖（仅打包机需要一次）**  
   - 若已装 Conda：打开 **Anaconda Prompt** 或 **命令提示符**，进入项目目录后执行：
   ```bat
   conda create -n file-mkdir python=3.10 -y
   conda activate file-mkdir
   conda install openpyxl -y
   pip install pyinstaller
   ```
   - 若只有 Python 无 Conda：
   ```bat
   pip install openpyxl pyinstaller
   ```

3. **执行打包**（在项目目录下）：
   ```bat
   build.bat
   ```
   或手动执行：
   ```bat
   pyinstaller --onefile --windowed --name "表格解析建目录" --hidden-import=openpyxl main.py
   ```

4. **取 exe**：打包完成后，在项目下的 `dist\` 文件夹里会有 **表格解析建目录.exe**。  
   - 把这个 exe（或整个 `dist` 文件夹）拷到别的 Windows 电脑即可使用。  
   - 对方电脑**不需要**安装 Python、Conda、openpyxl，双击 exe 即可打开 GUI。

### 云方案：本机没有 Windows 也能打出 exe

| 方案 | 说明 |
|------|------|
| **GitHub Actions**（推荐） | 代码推到 GitHub 后，在云端 Windows 自动打包，到 Actions 页下载 exe。免费、无需自备 Windows。 |
| **Gitee / 码云** | 使用 Gitee 的 CI（如 Gitee Go）配置 Windows 流水线，推送后构建并下载 exe。 |
| **云 Windows 桌面** | 购买按量计费的 Windows 云桌面（如腾讯云、阿里云、Azure），远程桌面进去按上面「在 Windows 上操作步骤」打包。 |
| **GitLab CI / Azure DevOps** | 在对应平台创建 Windows 流水线，执行与 `build.bat` 相同的安装和 PyInstaller 命令。 |

**使用 GitHub Actions（项目已含配置）**：

1. 在 GitHub 新建仓库，把本项目代码推上去（保留 `.github/workflows/build-windows-exe.yml`）。
2. 推送后打开仓库 **Actions** 页，选择刚运行的 “Build Windows exe” 任务。
3. 跑完后在页面底部 **Artifacts** 里下载 **表格解析建目录-exe**，解压即得 exe。

也可在 Actions 页点击 “Run workflow” 手动触发一次打包。

## 说明

- 客户编号：从「客户型号」中取**第一个空格+破折号前的数字**，如 `1116030298 - BG-...` → `1116030298`。
- 复制规则：在客户文件目录中，**名称包含该客户编号的压缩文件**（.zip、.rar、.7z、.tar.gz 等）和 **PDF** 会被复制到对应产品编号下的 `YG` 目录。若 YG 下已同时存在该客户的压缩文件和 PDF，则跳过复制。
