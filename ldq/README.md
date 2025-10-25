# 自动化操作与OCR识别工具

## 项目概述

这是一个功能强大的自动化操作与OCR识别工具，设计用于执行重复的屏幕操作任务。该工具可以通过识别屏幕上的特定文本内容和按钮位置，自动执行点击、输入等操作，极大地提高工作效率，减少重复性劳动。

## 功能特性

### 1. 屏幕区域OCR识别
- 支持选择多个屏幕区域进行实时OCR文本识别
- 可配置目标匹配文本，实现精准识别
- 支持文本标准化处理，提高识别准确性

### 2. 自动化操作
- 支持创建虚拟按钮并配置点击动作
- 可设置循环执行次数和执行间隔
- 支持设置停止条件，实现智能终止

### 3. 便捷控制
- 快捷键操作支持（F8开始/暂停，F9停止）
- 运行状态实时反馈
- 暂停-确认-恢复执行机制，防止误操作

### 4. 资源管理
- 自动清理临时文件和OCR缓存
- 安全释放系统资源
- 优雅的异常处理和错误提示

## 安装步骤

### 环境要求
- Python 3.8至3.10
- Windows操作系统（支持屏幕截图和UI操作）

### 安装依赖

1. 克隆或下载本项目到本地

2. 创建并激活虚拟环境（推荐）
   ```bash
   python -m venv env
   # Windows
   env\Scripts\activate
   # Linux/Mac
   source env/bin/activate
   ```

3. 安装必要的Python包（建议镜像安装）
   ```bash
   pip install -r requirements.txt
   ```

4. 安装特定版本的PaddlePaddle和PaddleX（OCR功能所需）
   ```bash
   pip install paddlex==2.0.0 paddlepaddle==2.4.2
   ```

## 使用方法

### 快速开始

1. 运行程序
   ```bash
   python ldq.py
   ```

2. 创建识别区域
   - 点击"添加识别区域"按钮
   - 设置目标匹配文本
   - 点击"选择区域"按钮，在屏幕上框选需要监控的区域

3. 创建虚拟按钮
   - 点击"添加按钮"按钮
   - 设置按钮显示文本和点击位置
   - 可选择是否启用键盘输入功能

4. 配置执行参数
   - 循环次数：设置自动操作的重复次数
   - 元素间间隔：设置识别区域之间的执行间隔
   - 按钮间间隔：设置按钮操作之间的执行间隔

5. 开始执行
   - 点击"开始执行"按钮或按下F8快捷键
   - 程序将按照配置的顺序执行所有操作

### 控制操作

- **暂停/恢复**：按下F8快捷键
- **停止执行**：按下F9快捷键或点击"停止当前循环"按钮
- **退出程序**：点击"退出程序"按钮

### 停止条件设置

1. 设置停止区域：选择一个屏幕区域用于监控停止条件
2. 配置停止文本：当该区域识别到指定文本时，程序将自动停止执行

## 项目结构

```
├── ldq.py           # 主程序文件
├── output/          # 输出目录（临时文件、截图等）
└── env/             # Python虚拟环境（可选）
```

## 贡献指南

我们欢迎社区贡献！如果您想为项目做出贡献，请遵循以下步骤：

1. Fork项目仓库
2. 创建您的特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交您的更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 开启一个Pull Request

### 代码规范

- 遵循PEP 8编码规范
- 添加适当的注释，特别是复杂逻辑
- 编写清晰的函数和变量名
- 避免使用硬编码的常量

## 问题排查

### 常见问题

1. **OCR识别失败**
   - 确保已正确安装paddlex==2.0.0和paddlepaddle==2.4.2
   - 检查识别区域是否合适，文本是否清晰可见

2. **快捷键不响应**
   - 确保程序窗口具有焦点
   - 检查是否有其他程序占用了相同的快捷键

3. **程序无响应**
   - 检查是否存在无限循环
   - 尝试增加执行间隔，避免CPU占用过高

## 开源许可证选择指南

### 开源许可证概述

开源许可证定义了软件的使用、修改和分发规则，为项目提供法律保护。选择合适的许可证对于项目的发展和社区贡献至关重要。

### 主要开源许可证比较

#### 1. MIT许可证
- **特点**：最宽松的许可证之一
- **限制**：要求保留版权声明和许可证文本
- **适用场景**：适合希望最大化代码复用的项目，企业友好
- **优点**：简单明了，使用门槛低，兼容性好

#### 2. Apache License 2.0
- **特点**：包含专利授权条款
- **限制**：要求保留版权声明和许可证文本，修改需注明
- **适用场景**：适合商业项目，特别是有专利考虑的项目
- **优点**：提供专利保护，防止专利侵权诉讼

#### 3. GNU GPL v3
- **特点**：传染性许可证
- **限制**：衍生作品必须使用相同许可证，源代码必须公开
- **适用场景**：强调软件自由，确保衍生作品也保持开源
- **优点**：保护软件自由，防止闭源商业化

#### 4. GNU LGPL v3
- **特点**：GPL的较弱版本，对链接库较宽松
- **限制**：修改库本身需要开源，但使用该库的应用可以闭源
- **适用场景**：开发通用库或组件，允许商业软件使用
- **优点**：平衡了开源贡献和商业使用需求

### 许可证选择建议

- **个人/小型项目**：MIT许可证是最常见的选择，简单且兼容性好
- **商业项目**：考虑Apache License 2.0提供的专利保护
- **自由软件项目**：如果希望保持软件自由，GPL系列是理想选择
- **库/框架开发**：LGPL允许更广泛的使用场景

### 在项目中添加许可证

1. **创建LICENSE文件**：在项目根目录创建一个名为`LICENSE`的文件
2. **选择许可证模板**：GitHub在创建仓库时提供常用许可证模板
3. **添加许可证头部**：在主要源代码文件顶部添加许可证头部注释

#### 许可证文件示例（MIT）

```
MIT License

Copyright (c) [year] [copyright holders]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

#### 代码文件头部示例

```python
# -*- coding: utf-8 -*-
"""
自动化操作与OCR识别工具

MIT License

Copyright (c) [year] [author name]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
"""
```

## 许可证

本项目采用MIT许可证开源 - 详情请查看 [LICENSE](LICENSE) 文件

## 联系方式

如有问题或建议，请联系：
- 邮箱：[your-email@example.com]
- GitHub：[your-github-username]

---

*祝您使用愉快！*