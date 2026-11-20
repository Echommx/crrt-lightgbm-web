# Dynamic Prediction of Antimicrobial Resistance Risk at Repeated Culture Events in Critically Ill Patients Receiving Continuous Renal Replacement Therapy

## 网页模型用途

这是一个仅用于研究和演示的本地 LightGBM 网页模型。模型根据患者基本信息、培养事件信息和培养前24小时动态观测，输出研究定义的两分类结果：

- `High risk`：预测分数达到研究阈值或更高。
- `Low risk`：预测分数低于研究阈值。

示例文件为模拟数据，不包含真实MIMIC-IV或珠江医院患者记录。

## 基本使用方法

1. 安装 `requirements-inference.txt` 中的推理依赖。
2. 运行 `python lightgbm_only_web.py --no-browser`。
3. 打开终端显示的本地地址，通常为 `http://127.0.0.1:8501`。
4. 点击 **Download Excel example** 下载 `simulated_input.xlsx`。
5. 保留工作表 `trend_input` 和第一行字段名，只修改模拟数据或替换为经过脱敏、授权的输入数据。
6. 确认动态观测位于T0前24小时内，保存为 `.xlsx`。
7. 点击 **Upload Excel (.xlsx)**，上传后点击 **Predict**。

模型输出是风险分层结果，不是耐药确诊结果，也不替代培养、药敏、感染科会诊或抗菌药物管理。

## 发布内容

`runtime_bundle.zip` 保存网页推理所需的最小运行模块。GitHub根目录不包含原始患者数据、训练数据、SHAP缓存、发表图表或其他模型。

新增样本应先完成结局审核、时间外验证、校准和性能评估，再按版本更新模型。样本增加不会自动保证模型性能提高。
