import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 页面基础配置
st.set_page_config(page_title="教师反馈完成率考核系统", layout="wide")
st.title("教师反馈完成率考核系统")
st.caption("普通班 / VIP班独立考核 · 上传Excel自动生成结果 · 支持一键导出")

# 侧边栏说明
with st.sidebar:
    st.header("📖 使用说明")
    st.markdown("""
    1. 上传Excel文件（普通班/VIP班均可，支持只传一个）
    2. 表格必须包含核心字段：
       - 项目/产品品类
       - 教师姓名
       - 是否进行班级学情反馈（Y/N）
    3. 系统自动只统计侧边栏名单里的老师
    4. 左侧可随时增删老师名单，自动重新计算
    """)
    st.markdown("---")
    st.subheader("📋 考核老师名单（可编辑）")
    default_names = """张洊维、姚蓓蓓、胡海枫、赵万铖、魏志如、吴恙、顾毅清云、杨安小、张玉玲"""
    name_input = st.text_area("修改名单，用顿号分隔", value=default_names, height=150)
    # 解析名单
    target_names = [n.strip() for n in name_input.replace("，", "、").split("、") if n.strip()]
    st.caption(f"当前共 {len(target_names)} 位老师")

    st.markdown("---")
    st.subheader("考核规则")
    st.markdown("""
    - 班级反馈完成率 = 班级反馈Y的课次 ÷ 总课次
    - **优秀**：完成率 ≥ 90%
    - **合格**：70% ≤ 完成率 ＜ 90%
    - **待改进**：完成率 ＜ 70%
    """)


def calculate_teacher_stats(df, target_names):
    """
    输入原始课次明细DataFrame，只统计目标名单里的老师，只算班级反馈
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

    # ========== 3. 只保留目标名单里的老师 ==========
    data = data[data["教师姓名"].isin(target_names)].reset_index(drop=True)
    if data.empty:
        st.warning("⚠️ 上传的表格里没有找到名单中的老师，请检查名字是否匹配")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # ========== 4. 处理班级学情反馈字段 ==========
    if "是否进行班级学情反馈" in data.columns:
        data["是否进行班级学情反馈"] = data["是否进行班级学情反馈"].astype(str).str.strip().str.upper()
        data["班级反馈完成"] = data["是否进行班级学情反馈"] == "Y"
    else:
        st.error("❌ 表格缺少「是否进行班级学情反馈」字段，请检查上传文件")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # ========== 5. 按 教师姓名 分组统计 ==========
    teacher_info = data.groupby("教师姓名").agg(
        负责项目=(project_col, lambda x: "、".join(sorted(x.dropna().unique())))
    ).reset_index()

    has_subject = "科目名称" in data.columns
    if has_subject:
        teacher_subjects = data.groupby("教师姓名")["科目名称"].apply(lambda x: "、".join(sorted(x.dropna().unique()))).reset_index()
        teacher_subjects.columns = ["教师姓名", "负责科目"]
        teacher_info = teacher_info.merge(teacher_subjects, on="教师姓名", how="left")

    stats = data.groupby("教师姓名").agg(
        总授课课次=("教师姓名", "count"),
        班级反馈完成课次=("班级反馈完成", "sum")
    ).reset_index()

    stats = stats.merge(teacher_info, on="教师姓名", how="left")

    # ========== 6. 计算完成率 ==========
    stats["班级反馈完成率"] = stats["班级反馈完成课次"] / stats["总授课课次"]
    stats["综合完成率"] = stats["班级反馈完成率"]  # 只算班级反馈

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
    stats_show["班级反馈完成率"] = stats_show["班级反馈完成率"].apply(lambda x: f"{x*100:.1f}%")
    stats_show["综合完成率"] = stats_show["综合完成率"].apply(lambda x: f"{x*100:.1f}%")

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


def generate_chart_image(normal_stats, vip_stats, target_names, title_text, subtitle_text):
    """生成「班课 vs VIP 按教师逐项对照」的反馈率图表（美观版），返回PNG字节"""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch
    from matplotlib import font_manager

    # ===== 中文字体：本地文件优先，否则自动下载（保证云端中文清晰） =====
    import tempfile, urllib.request
    font_candidates = ["SimHei.ttf", "NotoSansSC-Regular.otf"]
    font_path = None
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for cand in font_candidates:
        p = os.path.join(base_dir, cand)
        if os.path.exists(p):
            font_path = p
            break
    if font_path is None:
        # 仓库没有字体文件时，自动从网络下载SimHei（TTF，渲染稳定）
        try:
            tmp = os.path.join(tempfile.gettempdir(), "SimHei_auto.ttf")
            if not os.path.exists(tmp):
                urllib.request.urlretrieve(
                    "https://github.com/StellarCN/scp_zh/raw/master/fonts/SimHei.ttf",
                    tmp)
            font_path = tmp
        except Exception:
            font_path = None
    if font_path:
        font_manager.fontManager.addfont(font_path)
        font_name = font_manager.FontProperties(fname=font_path).get_name()
        plt.rcParams["font.sans-serif"] = [font_name, "Microsoft YaHei", "SimHei", "SimSun"]
    else:
        plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "SimSun"]
    plt.rcParams["axes.unicode_minus"] = False

    # 收集每个老师在班课和VIP的完成率（数值 0-1）
    rows = []
    for name in target_names:
        rate_n = None
        rate_v = None
        if normal_stats is not None and not normal_stats.empty:
            sub = normal_stats[normal_stats["教师姓名"] == name]
            if not sub.empty:
                rate_n = float(sub.iloc[0]["班级反馈完成率"])
        if vip_stats is not None and not vip_stats.empty:
            sub = vip_stats[vip_stats["教师姓名"] == name]
            if not sub.empty:
                rate_v = float(sub.iloc[0]["班级反馈完成率"])
        rows.append((name, rate_n, rate_v))

    # 按完成率从高到低排序（班课与VIP取非空值的平均作为排序依据）
    def sort_key(row):
        vals = [v for v in (row[1], row[2]) if v is not None]
        return sum(vals) / len(vals) if vals else 0.0
    rows.sort(key=sort_key, reverse=True)

    n = len(rows)
    fig_h = max(2.5, 2.1 + 0.72 * n)
    fig = plt.figure(figsize=(12, fig_h), dpi=150, facecolor="white")

    # ========== 标题区 ==========
    fig.text(0.5, 0.972, title_text, ha="center", va="center",
             fontsize=17, color="#1E293B")
    fig.text(0.5, 0.945, subtitle_text, ha="center", va="center",
             fontsize=11.5, color="#64748B")

    # ========== 布局坐标（fig相对坐标） ==========
    x_name_l, x_name_r = 0.030, 0.155        # 姓名列
    x_n_l, x_n_r = 0.185, 0.500              # 班课条形区
    x_v_l, x_v_r = 0.525, 0.840              # VIP条形区
    head_bottom = 0.902                      # 表头区底边
    head_h = 0.046                           # 表头高度
    body_top = head_bottom - 0.008           # 条形区顶边
    body_bottom = 0.085                      # 条形区底边
    body_h = body_top - body_bottom

    # ========== 表头（深蓝底白字 + 班课/VIP色块） ==========
    sw, sh = 0.014, 0.022                    # 表头内色块尺寸
    header_cells = [
        (x_name_l, x_name_r, "教师姓名", None),
        (x_n_l, x_n_r, "班课学情反馈率", "#3B82F6"),
        (x_v_l, x_v_r, "VIP学情反馈率", "#34D399"),
    ]
    for x0, x1, label, block_color in header_cells:
        rect = plt.Rectangle((x0, head_bottom), x1 - x0, head_h,
                             facecolor="#1F3864", edgecolor="none", zorder=5)
        fig.patches.append(rect)
        cx = (x0 + x1) / 2
        if block_color:
            # 色块嵌在表头文字左侧（深蓝底上的亮色块）
            fig.patches.append(plt.Rectangle(
                (cx - 0.062, head_bottom + (head_h - sh) / 2),
                sw, sh, facecolor=block_color, edgecolor="none", zorder=6))
            fig.text(cx + 0.005, head_bottom + head_h / 2, label,
                     ha="center", va="center", fontsize=11, color="white", zorder=6)
        else:
            fig.text(cx, head_bottom + head_h / 2, label,
                     ha="center", va="center", fontsize=11, color="white", zorder=6)

    # ========== 条形区 ==========
    ax_n = fig.add_axes([x_n_l, body_bottom, x_n_r - x_n_l, body_h])
    ax_v = fig.add_axes([x_v_l, body_bottom, x_v_r - x_v_l, body_h])

    y_pos = [1.0 - (i + 0.5) / n for i in range(n)]   # 第一行在顶部
    bar_h = 0.60 / n                                  # 条形高度随行数缩放

    # 配色：班课蓝 / VIP绿
    colors = {"n": "#2563EB", "v": "#10B981"}
    bgcolors = {"n": "#E2EAFD", "v": "#DCFCE7"}
    none_labels = {"n": "— 无班课", "v": "— 无VIP"}

    def draw_bars(ax, rates, color, bgcolor, none_label):
        # 行交替底色 + 满格浅色背景 + 圆角条形 + 数值
        for i, yi in enumerate(y_pos):
            # 交替行底色（仅奇数行）
            if i % 2 == 1:
                ax.axhspan((n - i - 1) / n, (n - i) / n,
                           facecolor="#F6F8FB", edgecolor="none", zorder=0)
        for yi, r in zip(y_pos, rates):
            # 满格浅色背景（0~1）
            rs = min(0.012, bar_h / 2)
            ax.add_patch(FancyBboxPatch(
                (0, yi - bar_h / 2), 1.0, bar_h,
                boxstyle=f"round,pad=0,rounding_size={rs}",
                linewidth=0, facecolor=bgcolor, zorder=2))
            # 实际比例条形
            if r is not None and r > 0:
                ax.add_patch(FancyBboxPatch(
                    (0, yi - bar_h / 2), r, bar_h,
                    boxstyle=f"round,pad=0,rounding_size={rs}",
                    linewidth=0, facecolor=color, zorder=3))
            # 数值标签
            if r is not None:
                ax.text(r + 0.015, yi, f"{r*100:.2f}%", va="center", ha="left",
                        fontsize=10.5, color="#1E293B", zorder=4)
            else:
                ax.text(0.02, yi, none_label, va="center", ha="left",
                        fontsize=9.5, color="#94A3B8", zorder=4)
        # 坐标细节
        ax.set_xlim(0, 1.30)
        ax.set_ylim(0, 1)
        ax.set_xticks([0, 1])
        ax.set_xticklabels(["0%", "100%"], fontsize=9.5, color="#64748B")
        ax.tick_params(axis="x", length=0, pad=4)
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        # 底部基线
        ax.axhline(0, color="#CBD5E1", linewidth=1.1, zorder=5)

    draw_bars(ax_n, [r[1] for r in rows], colors["n"], bgcolors["n"], none_labels["n"])
    draw_bars(ax_v, [r[2] for r in rows], colors["v"], bgcolors["v"], none_labels["v"])

    # ========== 教师姓名列（白底 + 交替底色 + 深色文字） ==========
    for i, (name, rn, rv) in enumerate(rows):
        if i % 2 == 1:
            fig.patches.append(plt.Rectangle(
                (x_name_l, body_bottom + body_h * (n - i - 1) / n),
                x_name_r - x_name_l, body_h / n,
                facecolor="#F6F8FB", edgecolor="none", zorder=1))
        fig_y = body_bottom + body_h * y_pos[i]
        fig.text((x_name_l + x_name_r) / 2, fig_y, name,
                 ha="center", va="center", fontsize=11.5,
                 color="#1E293B", zorder=4)

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return buf.getvalue()


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
            normal_stats, normal_show, normal_detail = calculate_teacher_stats(normal_raw, target_names)
            if not normal_stats.empty:
                has_normal = True
        except Exception as e:
            st.error(f"普通班文件读取失败：{str(e)}")

    # 处理VIP班
    if vip_file is not None:
        try:
            vip_raw = pd.read_excel(vip_file)
            vip_stats, vip_show, vip_detail = calculate_teacher_stats(vip_raw, target_names)
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
        if has_normal or has_vip:
            tab_names.append("📊 反馈率图表")
        if has_normal and has_vip:
            tab_names.append("综合汇总（普通+VIP）")
        if has_normal:
            tab_names.append("普通班考核排名")
        if has_vip:
            tab_names.append("VIP班考核排名")
        tab_names.append("待跟进教师清单")
        tab_names.append("课次明细溯源")

        tabs = st.tabs(tab_names)
        tab_i = 0

        # 图表标签页（班课 vs VIP 对照图，可下载）
        if has_normal or has_vip:
            with tabs[tab_i]:
                st.subheader("📊 学情反馈率图表（班课 vs VIP）")
                chart_title = st.text_input(
                    "图表标题",
                    value="贵阳新东方国际教育雅思阅写组学情反馈（9.1-10.16）",
                    key="chart_title"
                )
                chart_subtitle = st.text_input(
                    "副标题",
                    value="班课与VIP 按教师逐项对照",
                    key="chart_subtitle"
                )
                chart_img = generate_chart_image(
                    normal_stats if has_normal else None,
                    vip_stats if has_vip else None,
                    target_names,
                    chart_title,
                    chart_subtitle
                )
                st.image(chart_img, use_column_width=True)
                st.download_button(
                    label="📥 下载反馈率图表PNG",
                    data=chart_img,
                    file_name="学情反馈率图表.png",
                    mime="image/png"
                )
            tab_i += 1

        # 综合汇总标签页（只有同时上传两个表才显示）
        if has_normal and has_vip:
            with tabs[tab_i]:
                st.subheader("综合汇总：普通班+VIP班 整体反馈完成率")
                # 合并两个表的统计结果
                combined = pd.concat([
                    normal_stats[["教师姓名", "总授课课次", "班级反馈完成课次"]],
                    vip_stats[["教师姓名", "总授课课次", "班级反馈完成课次"]]
                ]).groupby("教师姓名").sum().reset_index()
                combined["班级反馈完成率"] = combined["班级反馈完成课次"] / combined["总授课课次"]
                
                def get_level(rate):
                    if rate >= 0.9:
                        return "优秀"
                    elif rate >= 0.7:
                        return "合格"
                    else:
                        return "待改进"
                combined["考核等级"] = combined["班级反馈完成率"].apply(get_level)
                combined = combined.sort_values("班级反馈完成率", ascending=False).reset_index(drop=True)
                
                # 格式化展示
                combined_show = combined.copy()
                combined_show["班级反馈完成率"] = combined_show["班级反馈完成率"].apply(lambda x: f"{x*100:.1f}%")
                st.dataframe(combined_show)
                
                # 同时展示各班型明细
                st.caption("各班型明细：")
                col_c1, col_c2 = st.columns(2)
                with col_c1:
                    st.write("普通班")
                    st.dataframe(normal_show)
                with col_c2:
                    st.write("VIP班")
                    st.dataframe(vip_show)
            tab_i += 1

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
