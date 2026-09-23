# 开发环境搭建

## 背景

系统 Python 由 Homebrew 托管 (`/usr/local/opt/python@3.13`),启用了 **PEP 668**,禁止全局 `pip install`,避免污染系统 Python。

## 方案

在项目根目录下创建独立虚拟环境 `.venv`,所有依赖安装在虚拟环境中。

```bash
# 一次性创建
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip

# 安装依赖
.venv/bin/python -m pip install pandas numpy openpyxl reportlab
```

## 已安装版本

| 包 | 版本 |
| --- | --- |
| Python | 3.13.15 |
| pandas | 3.0.6 |
| numpy | 2.5.3 |
| openpyxl | 3.1.5 |
| reportlab | 5.0.1 |

## 日常使用

```bash
# 激活虚拟环境(后续命令都走 venv 内的 python/pip)
source .venv/bin/activate

# 退出虚拟环境
deactivate

# 不激活也能直接调用
.venv/bin/python script.py
```

## 注意事项

- **不要**用 `pip install --break-system-packages` 绕过 PEP 668,会破坏 Homebrew Python。
- **不要**把 `.venv/` 提交到 Git,在 `.gitignore` 中忽略它。
- 新开 shell 需重新 `source .venv/bin/activate`,或直接用 `.venv/bin/python`。
