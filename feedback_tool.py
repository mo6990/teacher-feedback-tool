import streamlit as st
import pandas as pd
from io import BytesIO

# 页面基础配置
st.set_page_config(page_title="教师反馈完成率考核系统", layout="wide")
st.title("教师反馈完成率考核系统")
st.caption("普通班 / VIP班独立考核 · 上传Excel自动生成结果 · 支持一键导出")

# 侧边栏说明
with st.sidebar:
    st.header("📖 使用说明")
    st.markdown("""
    1. 上传两个Excel文件：普通班反馈表、VIP班反馈表
    2. 表格必须包含核心字段：
       - 项目
       - 教师姓名
       - 是否进行班级学情反馈（Y/N）
       - 是否进行个人学情反馈（格式：已反馈/总人数，如 5/6）
    3. 系统自动清洗重复表头、计算考核指标
    4. 支持按项目筛选、导出完整结果Excel
    """)
    st.markdown("---")
    st.subheader("考核规则")
    st.markdown("""
    - 班级反馈完成率 = 班级反馈Y的课次 ÷ 总课次
    - 个人反馈完成率 = 已反馈总人数 ÷ 应反馈总人数
    - 综合完成率 = 班级率×50% + 个人率×50%
    - **优秀**：综合 ≥ 90%
    - **合格**：70% ≤ 综合 ＜ 90%
    - **待改进**：综合 ＜ 70%
    """)


def calculate_teacher_stats(df):
    """
    输入原始课次明细DataFrame，自动适配普通班/VIP班格式，输出统计结果
    """
    data = df.copy()

    # 先统一列名去空格
    data.columns = data.columns.astype(str).str.strip()

    # ========== 1. 自动识别项目字段 ==========
    project_col = None
    if "项目" in data.columns:
        project_col = "项目"
    elif "产品品类" in data.columns:
        project_col = "产品品类"
    
    if project_col is None:
        st.error("❌ 表格缺少项目字段（「项目」或「产品品类」），请检查上传文件")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # ========== 2. 数据清洗：过滤空行、重复表头行 ==========
    data = data.dropna(subset=["教师姓名", project_col])
    data = data[data[project_col] != project_col]
    data = data.reset_index(drop=True)

    # ========== 3. 处理班级学情反馈字段 ==========
    if "是否进行班级学情反馈" in data.columns:
        data["是否进行班级学情反馈"] = data["是否进行班级学情反馈"].astype(str).str.strip().str.upper()
        data["班级反馈完成"] = data["是否进行班级学情反馈"] == "Y"
    else:
        st.error("❌ 表格缺少「是否进行班级学情反馈」字段，请检查上传文件")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # ========== 4. 处理个人学情反馈字段（自动识别格式） ==========
    if "是否进行个人学情反馈" not in data.columns:
        st.error("❌ 表格缺少「是否进行个人学情反馈」字段，请检查上传文件")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 自动判断是 Y/N 格式 还是 已反馈/总人数 格式
    sample_vals = data["是否进行个人学情反馈"].astype(str).head(20).tolist()
    has_slash = any("/" in str(v) for v in sample_vals)

    if has_slash:
        # 普通班格式：已反馈/总人数
        def split_feedback(x):
            try:
                x = str(x).strip()
                if "/" in x:
                    done, total = x.split("/", 1)
                    return int(done), int(total)
                return 0, 0
            except:
                return 0, 0
        data[["个人已反馈数", "个人总人数"]] = data["是否进行个人学情反馈"].apply(
            lambda x: pd.Series(split_feedback(x))
        )
    else:
        # VIP班格式：Y/N，每节课对应1个学员
        data["个人反馈完成"] = data["是否进行个人学情反馈"].astype(str).str.strip().str.upper() == "Y"
        data["个人已反馈数"] = data["个人反馈完成"].astype(int)
        data["个人总人数"] = 1  # 1对1课，每节课1个人

    # ========== 5. 按 教师姓名 分组统计（汇总所有项目、所有科目） ==========
    # 先统计每个老师带的项目和科目列表
    teacher_info = data.groupby("教师姓名").agg(
        负责项目=(project_col, lambda x: "、".join(sorted(x.dropna().unique())))
    ).reset_index()

    # 检查是否有科目名称列
    has_subject = "科目名称" in data.columns
    if has_subject:
        teacher_subjects = data.groupby("教师姓名")["科目名称"].apply(lambda x: "、".join(sorted(x.dropna().unique()))).reset_index()
        teacher_subjects.columns = ["教师姓名", "负责科目"]
        teacher_info = teacher_info.merge(teacher_subjects, on="教师姓名", how="left")

    stats = data.groupby("教师姓名").agg(
        总授课课次=("教师姓名", "count"),
        班级反馈完成课次=("班级反馈完成", "sum"),
        个人已反馈总数=("个人已反馈数", "sum"),
        个人应反馈总数=("个人总人数", "sum")
    ).reset_index()

    # 合并老师信息
    stats = stats.merge(teacher_info, on="教师姓名", how="left")

    # ========== 6. 计算各项完成率 ==========
    stats["班级反馈完成率"] = stats["班级反馈完成课次"] / stats["总授课课次"]
    stats["个人反馈完成率"] = stats.apply(
        lambda row: row["个人已反馈总数"] / row["个人应反馈总数"] if row["个人应反馈总数"] > 0 else 0,
        axis=1
    )
    stats["综合完成率"] = stats["班级反馈完成率"] * 0.5 + stats["个人反馈完成率"] * 0.5

    # ========== 7. 判定考核等级 ==========
    def get_level(rate):
        if rate >= 0.9:
            return "优秀"
        elif rate >= 0.7:
            return "合格"
        else:
            return "待改进"

    stats["考核等级"] = stats["综合完成率"].apply(get_level)

    # 按综合完成率降序排列
    stats = stats.sort_values("综合完成率", ascending=False).reset_index(drop=True)

    # 生成用于展示的百分比格式表
    stats_show = stats.copy()
    for col in ["班级反馈完成率", "个人反馈完成率", "综合完成率"]:
        stats_show[col] = stats_show[col].apply(lambda x: f"{x*100:.1f}%")

    return stats, stats_show, data


def generate_excel(normal_show, vip_show, normal_detail, vip_detail):
    """将所有结果打包为Excel文件，支持下载"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        normal_show.to_excel(writer, index=False, sheet_name="普通班考核结果")
        vip_show.to_excel(writer, index=False, sheet_name="VIP班考核结果")
        normal_detail.to_excel(writer, index=False, sheet_name="普通班课次明细")
        vip_detail.to_excel(writer, index=False, sheet_name="VIP班课次明细")
    return output.getvalue()


# ========== 页面主体：文件上传 ==========
col1, col2 = st.columns(2)
with col1:
    normal_file = st.file_uploader(
        "📤 上传普通班反馈表",
        type=["xlsx"],
        help="直接导出的原始Excel即可，无需提前清理重复表头"
    )
with col2:
    vip_file = st.file_uploader(
        "📤 上传VIP班反馈表",
        type=["xlsx"],
        help="字段格式与普通班保持一致"
    )

# ========== 文件上传后自动计算 ==========
if normal_file is not None or vip_file is not None:
    normal_stats, normal_show, normal_detail = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    vip_stats, vip_show, vip_detail = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    has_normal = False
    has_vip = False

    # 处理普通班
    if normal_file is not None:
        try:
            normal_raw = pd.read_excel(normal_file)
            normal_stats, normal_show, normal_detail = calculate_teacher_stats(normal_raw)
            if not normal_stats.empty:
                has_normal = True
        except Exception as e:
            st.error(f"普通班文件读取失败：{str(e)}")

    # 处理VIP班
    if vip_file is not None:
        try:
            vip_raw = pd.read_excel(vip_file)
            vip_stats, vip_show, vip_detail = calculate_teacher_stats(vip_raw)
            if not vip_stats.empty:
                has_vip = True
        except Exception as e:
            st.error(f"VIP班文件读取失败：{str(e)}")

    if has_normal or has_vip:
        # 整体概览看板
        st.subheader("📊 整体考核概览")
        cols = st.columns(4)
        col_idx = 0
        if has_normal:
            with cols[col_idx]:
                st.metric("普通班教师数", len(normal_stats))
            col_idx += 1
            with cols[col_idx]:
                normal_avg = normal_stats["综合完成率"].mean()
                st.metric("普通班综合均值", f"{normal_avg*100:.1f}%")
            col_idx += 1
        if has_vip:
            with cols[col_idx]:
                st.metric("VIP班教师数", len(vip_stats))
            col_idx += 1
            with cols[col_idx]:
                vip_avg = vip_stats["综合完成率"].mean()
                st.metric("VIP班综合均值", f"{vip_avg*100:.1f}%")
            col_idx += 1

        st.markdown("---")

        # 动态生成标签页
        tab_names = []
        if has_normal:
            tab_names.append("普通班考核排名")
        if has_vip:
            tab_names.append("VIP班考核排名")
        tab_names.append("待跟进教师清单")
        tab_names.append("课次明细溯源")

        tabs = st.tabs(tab_names)
        tab_i = 0

        if has_normal:
            with tabs[tab_i]:
                st.subheader("普通班教师反馈考核结果")
                # 提取所有涉及的科目
                all_subjects_normal = set()
                if "负责科目" in normal_stats.columns:
                    for s in normal_stats["负责科目"].dropna():
                        for subj in s.split("、"):
                            all_subjects_normal.add(subj)
                all_subjects_normal = sorted(list(all_subjects_normal))
                if all_subjects_normal:
                    select_normal = st.multiselect("按科目筛选", all_subjects_normal, default=all_subjects_normal, key="n_subj")
                    mask = normal_stats["负责科目"].apply(lambda x: any(subj in str(x) for subj in select_normal))
                    st.dataframe(normal_show[mask])
                else:
                    st.dataframe(normal_show)
            tab_i += 1

        if has_vip:
            with tabs[tab_i]:
                st.subheader("VIP班教师反馈考核结果")
                all_subjects_vip = set()
                if "负责科目" in vip_stats.columns:
                    for s in vip_stats["负责科目"].dropna():
                        for subj in s.split("、"):
                            all_subjects_vip.add(subj)
                all_subjects_vip = sorted(list(all_subjects_vip))
                if all_subjects_vip:
                    select_vip = st.multiselect("按科目筛选", all_subjects_vip, default=all_subjects_vip, key="v_subj")
                    mask = vip_stats["负责科目"].apply(lambda x: any(subj in str(x) for subj in select_vip))
                    st.dataframe(vip_show[mask])
                else:
                    st.dataframe(vip_show)
            tab_i += 1

        with tabs[tab_i]:
            st.subheader("⚠️ 待改进教师清单（综合完成率＜70%）")
            all_bad_list = []
            if has_normal:
                normal_bad = normal_show[normal_stats["综合完成率"] < 0.7].copy()
                normal_bad.insert(0, "班型", "普通班")
                all_bad_list.append(normal_bad)
            if has_vip:
                vip_bad = vip_show[vip_stats["综合完成率"] < 0.7].copy()
                vip_bad.insert(0, "班型", "VIP班")
                all_bad_list.append(vip_bad)
            all_bad = pd.concat(all_bad_list, ignore_index=True) if all_bad_list else pd.DataFrame()
            st.dataframe(all_bad)
            st.caption(f"共 {len(all_bad)} 位教师需重点跟进")
            tab_i += 1

        with tabs[tab_i]:
            if has_normal and has_vip:
                dt1, dt2 = st.tabs(["普通班明细", "VIP班明细"])
                with dt1:
                    st.dataframe(normal_detail)
                with dt2:
                    st.dataframe(vip_detail)
            elif has_normal:
                st.subheader("普通班课次明细")
                st.dataframe(normal_detail)
            elif has_vip:
                st.subheader("VIP班课次明细")
                st.dataframe(vip_detail)

        # 下载按钮
        st.markdown("---")
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            if has_normal:
                normal_show.to_excel(writer, index=False, sheet_name="普通班考核结果")
                normal_detail.to_excel(writer, index=False, sheet_name="普通班课次明细")
            if has_vip:
                vip_show.to_excel(writer, index=False, sheet_name="VIP班考核结果")
                vip_detail.to_excel(writer, index=False, sheet_name="VIP班课次明细")
        excel_file = output.getvalue()

        st.download_button(
            label="📥 下载考核结果Excel",
            data=excel_file,
            file_name="教师反馈完成率考核结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("👆 请上传至少一个Excel表格（普通班或VIP班均可），系统将自动完成计算")
