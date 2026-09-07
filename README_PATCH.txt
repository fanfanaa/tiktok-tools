合并补丁说明

本包同时包含：
1. SOP2 单选修复：radio 可正常选择，不再被页面记忆旧值覆盖。
2. SOP2 详细优缺点：逐条爆款/自己的作品都有优点、短板、为什么有效、优先改进；推荐组合有快速优缺点对比；深度对比展示双方优缺点；Excel同步扩展。
3. SOP2 3.8 -> 3.5 自动备用：3.8遇到429/503/UNAVAILABLE/HIGH DEMAND等临时错误会自动切换3.5。
4. 后台分析：SOP1 爆款解析/3方向/最终脚本，以及 SOP2 预分析/深度对比，都提交到后台线程池。切换页面不会中断，回来后继续显示进度或自动拿回结果。
5. 页面切换保留：沿用 workspace_service 的跨页状态/上传文件保留逻辑。

上传 GitHub main 覆盖/新增：
- background_service.py（新增）
- sop1_breakdown.py（覆盖）
- sop2_compare.py（覆盖）
- sop2_schema.py（覆盖）
- gemini_sop2.py（覆盖）
- workspace_service.py（覆盖）
- export_service.py（覆盖）

不需要改：
- app.py
- gemini_sop1.py（保留刚刚修好的 hotfix 版本）
- sop1_schema.py
- Supabase SQL / Secrets
- requirements.txt

注意：后台任务可跨页面继续，但如果 Streamlit Cloud 整个实例重启、休眠或重新部署，内存线程会被终止。正常页面切换不会中断。
