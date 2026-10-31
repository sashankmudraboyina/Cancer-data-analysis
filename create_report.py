import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def build_docx_report():
    print("Initializing document...")
    doc = docx.Document()

    # Page Margins: 0.85 inches for clean, modern look
    for s in doc.sections:
        s.top_margin = Inches(0.85)
        s.bottom_margin = Inches(0.85)
        s.left_margin = Inches(0.85)
        s.right_margin = Inches(0.85)

    # Palette Constants
    HEX_PRIMARY = "1B365D"      # Navy
    HEX_SECONDARY = "2B6CB0"    # Ocean / Slate Blue
    HEX_DARK = "2D3748"         # Charcoal Text
    HEX_LIGHT_BG = "F7FAFC"     # Off-white / Ice Gray
    HEX_CARD_BG = "EDF2F7"      # Muted Card Gray
    HEX_BORDER = "CBD5E0"       # Light Border
    HEX_ACCENT = "C53030"       # Crimson Accent
    HEX_HIGHLIGHT = "D69E2E"    # Amber / Gold Accent

    COLOR_PRIMARY = RGBColor(27, 54, 93)
    COLOR_SECONDARY = RGBColor(43, 108, 176)
    COLOR_DARK = RGBColor(45, 55, 72)
    COLOR_MUTED = RGBColor(113, 128, 150)

    # Configure Normal Style Font
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(10.5)
    font_normal.color.rgb = COLOR_DARK

    # Header and Footer Setup
    section = doc.sections[0]
    header = section.header
    p_head = header.paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_head = p_head.add_run("UNSUPERVISED MACHINE LEARNING FOR ONCOLOGICAL PROFILING | TECHNICAL RESEARCH REPORT")
    r_head.font.size = Pt(8.5)
    r_head.font.color.rgb = COLOR_MUTED

    footer = section.footer
    tbl_footer = footer.add_table(rows=1, cols=2, width=Inches(6.8))
    tbl_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_f_left = tbl_footer.cell(0, 0)
    cell_f_right = tbl_footer.cell(0, 1)
    
    p_f_l = cell_f_left.paragraphs[0]
    p_f_l.text = "Confidential | Healthcare Data Science & Informatics"
    p_f_l.runs[0].font.size = Pt(8.5)
    p_f_l.runs[0].font.color.rgb = COLOR_MUTED
    
    p_f_r = cell_f_right.paragraphs[0]
    p_f_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_f_r = p_f_r.add_run("Page ")
    r_f_r.font.size = Pt(8.5)
    r_f_r.font.color.rgb = COLOR_MUTED
    
    # Add page number
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r_f_r._r.append(fld1)
    r_f_r._r.append(instr)
    r_f_r._r.append(fld2)
    r_f_r._r.append(fld3)

    r_f_r2 = p_f_r.add_run(" of ")
    r_f_r2.font.size = Pt(8.5)
    r_f_r2.font.color.rgb = COLOR_MUTED
    fld4 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr2 = parse_xml(r'<w:instrText %s xml:space="preserve"> NUMPAGES </w:instrText>' % nsdecls('w'))
    fld5 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld6 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r_f_r2._r.append(fld4)
    r_f_r2._r.append(instr2)
    r_f_r2._r.append(fld5)
    r_f_r2._r.append(fld6)

    # -------------------------------------------------------------
    # Helper Functions
    # -------------------------------------------------------------
    def set_cell_background(cell, hex_color):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

    def add_callout(text, title=None, border_color=HEX_SECONDARY, bg_color=HEX_LIGHT_BG):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.8)
        set_cell_background(cell, bg_color)
        set_cell_margins(cell, top=140, bottom=140, left=200, right=160)
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        cell._tc.get_or_add_tcPr().append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        if title:
            rt = p.add_run(f"{title}\n")
            rt.font.bold = True
            rt.font.size = Pt(10.5)
            rt.font.color.rgb = RGBColor.from_string(border_color)
        rx = p.add_run(text)
        rx.font.size = Pt(10)
        rx.font.italic = True
        rx.font.color.rgb = COLOR_DARK
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)

    def add_stat_cards(stats):
        cols = len(stats)
        tbl = doc.add_table(rows=1, cols=cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        w = Inches(6.8 / cols)
        for idx, (num, lbl) in enumerate(stats):
            c = tbl.cell(0, idx)
            c.width = w
            set_cell_background(c, "F8FAFC")
            set_cell_margins(c, top=160, bottom=160, left=100, right=100)
            b = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="6" w:space="0" w:color="E2E8F0"/>'
                f'<w:left w:val="single" w:sz="6" w:space="0" w:color="E2E8F0"/>'
                f'<w:bottom w:val="single" w:sz="24" w:space="0" w:color="{HEX_PRIMARY}"/>'
                f'<w:right w:val="single" w:sz="6" w:space="0" w:color="E2E8F0"/>'
                f'</w:tcBorders>'
            )
            c._tc.get_or_add_tcPr().append(b)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            rn = p.add_run(num + "\n")
            rn.font.bold = True
            rn.font.size = Pt(22)
            rn.font.color.rgb = COLOR_PRIMARY
            rl = p.add_run(lbl)
            rl.font.size = Pt(8.5)
            rl.font.bold = True
            rl.font.color.rgb = COLOR_MUTED
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(6)

    def add_fig(img_path, caption_text, figure_num, width_in_inches=6.4):
        if not os.path.exists(img_path):
            print(f"Warning: Missing {img_path}")
            return
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(3)
        r = p_img.add_run()
        r.add_picture(img_path, width=Inches(width_in_inches))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(10)
        rf = p_cap.add_run(f"Figure {figure_num}: ")
        rf.font.bold = True
        rf.font.size = Pt(9.5)
        rf.font.color.rgb = COLOR_PRIMARY
        rc = p_cap.add_run(caption_text)
        rc.font.size = Pt(9.5)
        rc.font.italic = True
        rc.font.color.rgb = COLOR_DARK

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(20)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.bold = True
        r.font.size = Pt(16)
        r.font.color.rgb = COLOR_PRIMARY
        # Subtle horizontal divider line under H1
        p_line = doc.add_paragraph()
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(6)
        r_l = p_line.add_run("―" * 58)
        r_l.font.color.rgb = RGBColor(203, 213, 224)
        r_l.font.size = Pt(8)

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = COLOR_SECONDARY

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = COLOR_DARK

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text)
        return p

    def add_bullet(text, prefix="• "):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r_b = p.add_run(prefix)
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_SECONDARY
        r_t = p.add_run(text)
        return p

    # -------------------------------------------------------------
    # 1. TITLE & COVER HEADER
    # -------------------------------------------------------------
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(12)
    p_pre.paragraph_format.space_after = Pt(4)
    r_pre = p_pre.add_run("RESEARCH MONOGRAPH & BENCHMARK REPORT")
    r_pre.font.bold = True
    r_pre.font.size = Pt(10)
    r_pre.font.color.rgb = COLOR_SECONDARY

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(8)
    r_t = p_title.add_run("Comparative Clustering Analysis and Algorithmic Benchmarking in Oncology")
    r_t.font.bold = True
    r_t.font.size = Pt(24)
    r_t.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(16)
    r_s = p_sub.add_run("A Multi-Scale Unsupervised Learning Investigation Across Global Epidemiological Cancer Mortality (204 Sovereign Nations) and Microscopic Cellular Pathology (569 Biopsies) Across 16 Clustering Architectures")
    r_s.font.size = Pt(12)
    r_s.font.color.rgb = COLOR_MUTED

    # Meta card table
    tbl_meta = doc.add_table(rows=2, cols=4)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    meta_w = Inches(6.8 / 4)
    headers_meta = ["Domain", "Methodology", "Algorithms", "Benchmark Date"]
    values_meta = ["Bioinformatics & Oncology", "Dual-Scale Unsupervised", "16 ML Models (6 Families)", "October 2026"]
    
    for i in range(4):
        c0 = tbl_meta.cell(0, i)
        c1 = tbl_meta.cell(1, i)
        c0.width = meta_w
        c1.width = meta_w
        set_cell_background(c0, HEX_PRIMARY)
        set_cell_background(c1, "F8FAFC")
        set_cell_margins(c0, top=80, bottom=80, left=100, right=100)
        set_cell_margins(c1, top=100, bottom=100, left=100, right=100)
        
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(headers_meta[i])
        r0.font.bold = True
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = RGBColor(255, 255, 255)
        
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(values_meta[i])
        r1.font.size = Pt(9)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_DARK

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(8)
    p_sp.paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # 2. EXECUTIVE SUMMARY
    # -------------------------------------------------------------
    add_h1("1. Executive Summary")

    add_p("Unsupervised machine learning provides an indispensable set of mathematical tools for exploring unlabelled clinical datasets, uncovering hidden disease archetypes, stratifying patient cohorts, and discovering fundamental epidemiological groupings. This report details an exhaustive, multi-scale clustering investigation designed to benchmark 16 distinct clustering algorithms across six foundational algorithmic families. The analysis spans two diametrically opposed observational scales in oncology:")

    add_bullet("Macro-Epidemiological Scale: 204 sovereign nations evaluated across 22 age-standardized cancer mortality rates per 100,000 population (World Health Organization / Institute for Health Metrics and Evaluation Global Burden of Disease study, 2019 cross-section). This dimension investigates how structural development, environmental exposures, lifestyle factors, and healthcare systems shape national disease burdens.")
    add_bullet("Micro-Pathological Cellular Scale: 569 fine-needle aspirate (FNA) biopsies from the Wisconsin Diagnostic Breast Cancer (WDBC) cohort, parameterized into 30 high-dimensional continuous morphological descriptors of cell nuclei. This dimension tests whether purely unsupervised algorithms can autonomously re-discover the biological distinction between benign tissue and malignant tumors without supervision.")

    # Executive Stat KPI cards
    add_stat_cards([
        ("16", "Algorithms Evaluated"),
        ("204", "Nations Profiled"),
        ("569", "Biopsies Partitioned"),
        ("0.774", "Peak Rand Index (GMM)")
    ])

    add_callout(
        "CORE EMPIRICAL TAKEAWAY: Gaussian Mixture Models (GMM) and Graph Spectral Clustering achieve superior performance in delineating cellular malignancy boundaries (ARI = 0.7740 and 0.7608, respectively), significantly outperforming standard K-Means (ARI = 0.6536). Conversely, density-based algorithms (DBSCAN, HDBSCAN, OPTICS) suffer catastrophic failure in continuous high-dimensional clinical spaces due to the curse of dimensionality and uniform density dissipation. At the global scale, cancer mortality cleaves into four sharply demarcated geopolitical archetypes driven by the epidemiological transition.",
        "EXECUTIVE STRATEGIC SUMMARY",
        HEX_PRIMARY,
        "F0F4F8"
    )

    add_h2("1.1 Key Scientific and Strategic Findings")
    add_bullet("Superiority of Probabilistic & Manifold Geometries: Gaussian Mixture Models (EM with unconstrained full covariance) accurately capture the elliptical, overlapping distributions of nuclear atypia, providing calibrated soft assignment probabilities that mirror biological tumor grading uncertainty.")
    add_bullet("The Curse of Dimensionality in Clinical Density Clustering: Density algorithms (DBSCAN, HDBSCAN, OPTICS) failed completely on the cellular dataset, labeling between 5.6% and 100% of samples as unclassifiable noise. In 30-dimensional space, the ratio between the distance to the nearest neighbor and farthest neighbor approaches unity, obliterating the density gradients required by epsilon-neighborhood methods.")
    add_bullet("The Chaining Vulnerability of Single Linkage: Agglomerative clustering with single linkage registered an artificially inflated Silhouette score (0.6607 on Wisconsin, 0.4527 on Global) while achieving a near-zero Adjusted Rand Index (ARI = 0.0048). This mathematical artifact arises when a single dense mega-cluster absorbs 99% of points, leaving isolated outliers as singletons.")
    add_bullet("Four Global Epidemiological Archetypes: Global cancer mortality cleanly segregates into: (1) Emerging Nations characterized by stomach and esophageal cancer; (2) Sub-Saharan & Island Nations burdened by cervical and prostate cancer; (3) A distinct Endemic Island Outlier (Tonga) marked by extreme testicular and liver mortality; and (4) Industrialized High-SDI Nations dominated by lung, colorectal, pancreatic, kidney, and CNS cancers.")

    # -------------------------------------------------------------
    # 3. TAXONOMY & MATHEMATICAL FOUNDATIONS
    # -------------------------------------------------------------
    add_h1("2. Algorithmic Taxonomy & Theoretical Foundations")

    add_p("Clustering algorithms differ substantially in their underlying mathematical objectives, distance metrics, geometric cluster shape assumptions, and computational scaling profiles. To establish a rigorous comparative taxonomy, the 16 algorithms investigated in this study were categorized into six formal algorithmic families:")

    add_h2("2.1 Centroid-Based & Partitioning Algorithms")
    add_p("Centroid methods partition data space into Voronoi cells by iteratively updating cluster centroids to minimize within-cluster sum of squares (inertia):")
    add_bullet("1. K-Means (Lloyd's Algorithm): Iteratively alternates between assigning each sample to its nearest centroid in Euclidean space and recomputing centroids as the arithmetic mean of assigned points. Assumes convex, isotropic (spherical) clusters of roughly equal variance.")
    add_bullet("2. Mini-Batch K-Means (Sculley, 2010): Uses stochastic mini-batches of subsampled data points to perform incremental gradient updates to cluster centers, drastically reducing memory bandwidth and cache misses while retaining convergence fidelity.")
    add_bullet("3. Bisecting K-Means: A divisive hierarchical variant that starts with all points in a single cluster and iteratively bisects existing clusters using standard 2-means until the target k is reached. This divisive tree structure provides greater robustness against sub-optimal local minima.")

    add_h2("2.2 Hierarchical Agglomerative Clustering")
    add_p("Hierarchical agglomeration establishes a nested bottom-up dendrogram by iteratively merging the pair of clusters that minimizes a specified linkage distance criterion:")
    add_bullet("4. Ward Linkage: Minimizes the total within-cluster variance increase (sum of squared deviations from the merged mean) at each agglomeration step. Creates compact, balanced spherical clusters.")
    add_bullet("5. Complete Linkage: Merges clusters based on the maximum pairwise distance between points in each cluster, prioritizing compact, tightly bounded clusters with equal diameters.")
    add_bullet("6. Average Linkage (UPGMA): Merges clusters based on the average pairwise distance across all constituent elements, serving as a compromise between Ward and single linkage.")
    add_bullet("7. Single Linkage: Merges clusters based on the minimum pairwise distance between constituent points. Highly sensitive to the 'chaining effect', where thin bridges of noise merge disparate structures.")

    add_h2("2.3 Density-Based Algorithms")
    add_p("Density methods discover arbitrarily shaped clusters by identifying spatial regions with a point density exceeding user-defined thresholds, designating low-density regions as noise:")
    add_bullet("8. DBSCAN (Ester et al., 1996): Connects core points that possess at least min_samples neighbors within an epsilon radius. Points neither core nor border are flagged as noise (-1). Assumes uniform spatial density across clusters.")
    add_bullet("9. HDBSCAN (Campello et al., 2013): Converts DBSCAN into a hierarchical framework over all possible epsilon values, constructing a mutual reachability distance graph and extracting stable clusters using an excess-of-mass persistence criterion.")
    add_bullet("10. OPTICS (Ankerst et al., 1999): Creates an ordered reachability graph storing the core distance and reachability distance for each point, allowing interactive extraction of clusters across varying density levels.")

    add_h2("2.4 Centroid Shift / Mode-Seeking")
    add_bullet("11. Mean Shift (Comaniciu & Meer, 2002): A non-parametric mode-seeking algorithm that models feature space as an empirical probability density function estimated via a kernel (e.g., Gaussian RBF). Data points perform gradient ascent iteratively along the density gradient until convergence at local modes.")

    add_h2("2.5 Graph & Manifold Spectral Clustering")
    add_bullet("12. Spectral Clustering (Ng, Jordan, & Weiss, 2001): Maps data points into nodes of an affinity graph (e.g., k-nearest neighbors or RBF kernel), constructs the Normalized Graph Laplacian L = I - D^(-1/2) W D^(-1/2), projects the top eigenvectors into a low-dimensional manifold subspace, and clusters the embedded points via K-Means. Capable of disentangling highly non-linear, non-convex manifolds.")

    add_h2("2.6 Exemplar-Based & Tree-Subspace Clustering")
    add_bullet("13. Affinity Propagation (Frey & Dueck, 2007): Identifies representative exemplar points by passing messages of 'responsibility' r(i,k) and 'availability' a(i,k) between data points across iterations until convergence, automatically determining cluster counts based on input preferences.")
    add_bullet("14. BIRCH (Zhang, Ramakrishnan, & Livny, 1996): Constructs an in-memory Clustering Feature (CF) tree summarizing multi-dimensional data via zero, first, and second order statistical moments. Rapidly compresses massive datasets before running hierarchical clustering on leaf nodes.")

    add_h2("2.7 Probabilistic & Bayesian Mixture Models")
    add_bullet("15. Gaussian Mixture Models [GMM] (Dempster, Laird, & Rubin, 1977): Generative probabilistic model assuming each data point is drawn from a mixture of k Gaussian distributions with parameterized mean vectors and full covariance matrices. Optimized via Expectation-Maximization (EM), yielding calibrated posterior probabilities of cluster membership.")
    add_bullet("16. Bayesian Gaussian Mixture Models [DPGMM]: Extends GMM with a Dirichlet Process prior over mixture weights. Uses variational inference to infer cluster allocations while automatically penalizing unnecessary mixture components, driving inactive cluster weights toward zero.")

    # Comparative Architecture Table
    add_h3("Table 1: Algorithmic Architecture & Theoretical Comparison Matrix")
    tbl_tax = doc.add_table(rows=17, cols=6)
    tbl_tax.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_tax.autofit = False
    
    headers_tax = ["Algorithm", "Family", "Geometry / Shape", "Complexity (Time)", "Sensitivity / Failure Mode", "k Strategy"]
    tax_widths = [Inches(1.3), Inches(1.1), Inches(1.1), Inches(1.0), Inches(1.4), Inches(0.9)]
    
    for j in range(6):
        c = tbl_tax.cell(0, j)
        c.width = tax_widths[j]
        set_cell_background(c, HEX_PRIMARY)
        set_cell_margins(c, top=80, bottom=80, left=80, right=80)
        p = c.paragraphs[0]
        r = p.add_run(headers_tax[j])
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    tax_rows = [
        ("1. K-Means", "Centroid", "Convex / Spherical", "O(k · n · d · i)", "Outliers, unequal variance", "Pre-specified"),
        ("2. Mini-Batch K-Means", "Centroid", "Convex / Spherical", "O(b · k · d · i)", "Slightly higher variance", "Pre-specified"),
        ("3. Bisecting K-Means", "Centroid", "Divisive Spherical", "O(n · d · log k)", "Local minima in bisections", "Pre-specified"),
        ("4. Agglomerative (Ward)", "Hierarchical", "Compact Spherical", "O(n² · log n)", "Memory scale on large n", "Dendrogram Cut"),
        ("5. Agglomerative (Complete)", "Hierarchical", "Compact / Equal size", "O(n² · log n)", "Outlier sensitivity", "Dendrogram Cut"),
        ("6. Agglomerative (Average)", "Hierarchical", "Intermediate", "O(n² · log n)", "Variable density distortion", "Dendrogram Cut"),
        ("7. Agglomerative (Single)", "Hierarchical", "Arbitrary non-convex", "O(n²)", "Severe chaining vulnerability", "Dendrogram Cut"),
        ("8. DBSCAN", "Density", "Arbitrary Manifold", "O(n · log n) to O(n²)", "High-d uniform density collapse", "Epsilon threshold"),
        ("9. HDBSCAN", "Density", "Variable Arbitrary", "O(n²)", "Curse of dimensionality (>20d)", "Min cluster size"),
        ("10. OPTICS", "Density", "Multi-density", "O(n²)", "High dimensional noise bias", "Reachability plot"),
        ("11. Mean Shift", "Mode-Seeking", "Arbitrary Mode", "O(i · n²)", "Bandwidth estimation failure", "Automatic"),
        ("12. Spectral Clustering", "Graph / Manifold", "Non-convex Manifold", "O(n³)", "Graph affinity parameter tuning", "Pre-specified"),
        ("13. Affinity Propagation", "Exemplar", "Exemplar-based", "O(n² · i)", "Over-partitioning, memory", "Preference value"),
        ("14. BIRCH", "Tree-Subspace", "Spherical Subclusters", "O(n · d)", "Non-spherical subclusters", "Threshold / Branch"),
        ("15. Gaussian Mixture (GMM)", "Probabilistic", "Unconstrained Elliptical", "O(i · n · d³)", "Singular covariance matrices", "Pre-specified"),
        ("16. Bayesian GMM (DPGMM)", "Bayesian", "Unconstrained Elliptical", "O(i · n · d³)", "Prior hyperparameter drift", "Pruned DP Prior")
    ]
    
    for row_idx, data in enumerate(tax_rows):
        bg = HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
        for col_idx in range(6):
            cell = tbl_tax.cell(row_idx + 1, col_idx)
            cell.width = tax_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            r = p.add_run(data[col_idx])
            r.font.size = Pt(8)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = COLOR_PRIMARY
            else:
                r.font.color.rgb = COLOR_DARK

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(6)
    p_sp.paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # 4. EXPERIMENTAL DESIGN & EXPLORATORY ANALYSIS
    # -------------------------------------------------------------
    add_h1("3. Experimental Design, Data Architecture, & Exploratory Analysis")

    add_p("To test the generalizability of clustering paradigms across differing data topologies, two distinct biomedical and public health datasets were constructed, standardized, and benchmarked:")

    add_h2("3.1 Dataset Architecture")
    add_bullet("Global Cancer Mortality Dataset (Macro): Sourced from the Institute for Health Metrics and Evaluation (IHME) Global Burden of Disease study (via Our World in Data). Filtered for sovereign nations with valid ISO-3 country codes for the reporting year 2019. The dataset comprises N = 204 sovereign nations across D = 22 cancer mortality rates, measured as age-standardized deaths per 100,000 population to control for national demographic age distributions.")
    add_bullet("Wisconsin Diagnostic Breast Cancer Dataset (Micro): Acquired from the UCI Machine Learning Repository and digitized by Dr. William H. Wolberg, W. Nick Street, and Olvi L. Mangasarian. Comprises N = 569 fine-needle aspirate (FNA) biopsy samples of human breast masses across D = 30 continuous nuclear morphological features. Features represent the mean, standard error, and 'worst' (mean of three largest values) of ten fundamental geometric properties: radius, texture, perimeter, area, smoothness, compactness, concavity, concave points, symmetry, and fractal dimension. Ground truth histopathological diagnoses classify 357 tumors as Benign and 212 as Malignant.")

    add_h2("3.2 Preprocessing, Standardization, & Principal Component Analysis")
    add_p("Continuous feature dimensions vary by multiple orders of magnitude (e.g., breast tumor area ranges from 143.5 to 2,501.0 mm², whereas fractal dimension spans 0.05 to 0.10). In the absence of standardization, distance-based metrics would be dominated by features with large absolute variances. Standard Z-score feature scaling was applied independently across both feature spaces: z = (x - μ) / σ, yielding zero mean and unit variance for every feature.")

    add_p("Dimensionality reduction via Principal Component Analysis (PCA) was conducted to quantify the intrinsic dimensionality and variance distribution:")
    add_bullet("Global Cancer Dataset: Principal Component 1 (PC1) accounts for 22.71% of total variance, and Principal Component 2 (PC2) explains 13.74%, yielding a cumulative 2-component explained variance of 36.45%. Five components capture 60.10% of total variance, reflecting the complex, multi-factorial etiology of distinct cancer types.")
    add_bullet("Wisconsin WDBC Dataset: PC1 captures 44.27% of variance, and PC2 captures 18.97%, achieving a cumulative 2-component variance of 63.24%. Five components capture 84.76% of total variance. This high concentration in PC1 indicates strong biological co-variation among cell size, perimeter, concavity, and malignancy.")

    add_h2("3.3 Optimal Number of Clusters (k) Determination")
    add_p("To establish the optimal cluster parameter k for partition-based algorithms, within-cluster inertia (Elbow method) and mean Silhouette coefficients were evaluated across k ∈ [2, 10]:")

    add_fig("plots/01_optimal_k_global_cancer.png", "Elbow Method (Inertia vs k) and Silhouette Coefficient curves for the Global Cancer Mortality dataset. The silhouette peak occurs at k=2, but k=4 yields the definitive epidemiological elbow and captures clinically distinct geopolitical profiles.", 1)

    add_fig("plots/02_optimal_k_wisconsin.png", "Elbow Method and Silhouette Coefficient curves for the Wisconsin Diagnostic Breast Cancer dataset. Silhouette score peaks sharply at k=2 (Score = 0.3434), aligning perfectly with the underlying binary biological ground truth (Malignant vs Benign).", 2)

    add_p("For the Global Cancer dataset, while k = 2 produces a mathematical silhouette peak, it merely bisects nations into broad high vs low mortality groups. At k = 4, the curve exhibits a pronounced elbow where the marginal reduction in inertia levels off, while preserving four distinct and actionable epidemiological archetypes. For the Wisconsin dataset, k = 2 represents both the empirical silhouette maximum and the biological ground truth.")

    add_h2("3.4 Inter-Cancer Correlation Structure & Hierarchical Taxonomy")
    add_p("To examine how specific oncological pathologies co-occur across global populations, a 22 × 22 Pearson correlation matrix was computed across age-standardized mortality rates:")

    add_fig("plots/04_cancer_types_correlation_matrix.png", "Global Correlation Heatmap of 22 Age-Standardized Cancer Mortality Rates (2019). Two powerful opposing etiological clusters emerge: an Industrialized/Lifestyle nexus and an Infectious/Screening-Deficit nexus.", 3)

    add_callout(
        "ETIOLOGICAL CO-OCCURRENCE HIGHLIGHT: Kidney cancer exhibits strong positive correlations with Pancreatic cancer (r = 0.774) and Colorectal cancer (r = 0.715). Colorectal cancer is tightly linked with Lung cancer (r = 0.680) and Ovarian cancer (r = 0.649). Conversely, Cervical cancer correlates negatively with Brain/CNS cancer (r = -0.552) and Kidney cancer (r = -0.540). This division demarcates non-communicable, metabolic/lifestyle malignancies in wealthy countries from infection-driven, screening-deficit malignancies in developing countries.",
        "EPIDEMIOLOGICAL CORRELATION ANALYSIS",
        HEX_SECONDARY,
        HEX_LIGHT_BG
    )

    add_fig("plots/03_dendrogram_global_cancer.png", "Hierarchical Clustering Dendrogram (Ward's Minimum Variance Linkage) over 204 sovereign nations labeled by ISO-3 codes. The hierarchical tree reveals major continental and socio-demographic bifurcations.", 4)

    # -------------------------------------------------------------
    # 5. CROSS-ALGORITHMIC BENCHMARK
    # -------------------------------------------------------------
    add_h1("4. Cross-Algorithmic Benchmark & Comparative Evaluation")

    add_p("To comprehensively benchmark all 16 algorithms, an empirical evaluation framework was deployed incorporating both internal validation metrics (measuring cluster cohesion and separation without labels) and external ground truth validation metrics (measuring concordance against histological biopsy diagnoses):")

    add_bullet("Silhouette Score: Measures how similar an object is to its own cluster compared to neighboring clusters, ranging from -1 (poor) to +1 (dense, well-separated).")
    add_bullet("Davies-Bouldin Index (DB): Ratio of within-cluster scatter to between-cluster separation. Lower scores indicate superior clustering.")
    add_bullet("Calinski-Harabasz Index (CH): Variance ratio criterion measuring between-cluster dispersion over within-cluster dispersion. Higher scores reflect more compact, isolated clusters.")
    add_bullet("Adjusted Rand Index (ARI): Measures pair-wise agreement between cluster assignments and ground truth labels, adjusted for chance (1.0 indicates perfect classification; 0.0 indicates chance-level agreement).")
    add_bullet("Normalized Mutual Information (NMI): Quantifies shared information between cluster assignments and ground truth, normalized between 0.0 and 1.0.")
    add_bullet("V-Measure: Harmonic mean of homogeneity (each cluster contains only members of a single class) and completeness (all members of a class are assigned to the same cluster).")

    add_h2("4.1 Global Cancer Mortality Clustering Benchmark (204 Nations)")
    add_p("Table 2 presents the quantitative benchmarking results across all 16 algorithms applied to the standardized 22-dimensional global cancer mortality dataset at target k = 4:")

    # Table 2: Global Benchmark
    tbl_b_g = doc.add_table(rows=17, cols=6)
    tbl_b_g.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_b_g.autofit = False
    
    headers_bg = ["Algorithm", "Clusters (k)", "Noise %", "Silhouette", "Davies-Bouldin", "Calinski-Harabasz"]
    bg_widths = [Inches(1.8), Inches(0.9), Inches(0.9), Inches(1.1), Inches(1.1), Inches(1.0)]
    
    for j in range(6):
        c = tbl_b_g.cell(0, j)
        c.width = bg_widths[j]
        set_cell_background(c, HEX_PRIMARY)
        set_cell_margins(c, top=80, bottom=80, left=80, right=80)
        p = c.paragraphs[0]
        r = p.add_run(headers_bg[j])
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    global_data = [
        ("1. K-Means", "4", "0.0%", "0.1924", "1.5253", "28.42"),
        ("2. Mini-Batch K-Means", "4", "0.0%", "0.1270", "2.1170", "25.93"),
        ("3. Bisecting K-Means", "4", "0.0%", "0.1194", "2.1141", "26.83"),
        ("4. Agglomerative (Ward)", "4", "0.0%", "0.1403", "1.9992", "25.46"),
        ("5. Agglomerative (Complete)", "4", "0.0%", "0.4271*", "0.6457*", "9.93"),
        ("6. Agglomerative (Average)", "4", "0.0%", "0.4271*", "0.6457*", "9.93"),
        ("7. Agglomerative (Single)", "4", "0.0%", "0.4527*", "0.3470*", "9.03"),
        ("8. DBSCAN", "2", "14.7%", "0.2280", "0.8618", "8.31"),
        ("9. HDBSCAN", "5", "69.1%", "0.4121", "1.0065", "37.18"),
        ("10. OPTICS", "12", "62.7%", "0.3984", "0.9143", "31.98"),
        ("11. Mean Shift", "6", "0.0%", "0.2319", "0.8728", "8.49"),
        ("12. Spectral Clustering", "4", "0.0%", "0.1477", "1.9619", "25.92"),
        ("13. Affinity Propagation", "30", "0.0%", "0.1802", "1.0048", "17.88"),
        ("14. BIRCH", "4", "0.0%", "0.1403", "1.9992", "25.46"),
        ("15. Gaussian Mixture (GMM)", "4", "0.0%", "0.1833", "1.8360", "25.86"),
        ("16. Bayesian GMM (DPGMM)", "4", "0.0%", "0.1794", "1.8601", "25.51")
    ]

    for row_idx, row in enumerate(global_data):
        bg = HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
        for col_idx in range(6):
            cell = tbl_b_g.cell(row_idx + 1, col_idx)
            cell.width = bg_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            if col_idx in [1, 2, 3, 4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(row[col_idx])
            r.font.size = Pt(8.5)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = COLOR_PRIMARY
            else:
                r.font.color.rgb = COLOR_DARK

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(2)
    p_note.paragraph_format.space_after = Pt(8)
    r_n = p_note.add_run("*Note: High Silhouette scores for Single, Complete, and Average linkage reflect extreme cluster imbalance (chaining artifact) where >98% of points reside in one cluster and 3 single outlier points form tiny peripheral clusters.")
    r_n.font.size = Pt(8)
    r_n.font.italic = True
    r_n.font.color.rgb = COLOR_MUTED

    add_h2("4.2 Wisconsin Diagnostic Breast Cancer Diagnostic Benchmark (569 Tumors)")
    add_p("Table 3 presents the quantitative benchmarking results across all 16 algorithms evaluated against the histological ground truth diagnoses (357 Benign vs 212 Malignant) at target k = 2:")

    # Table 3: Wisconsin Benchmark
    tbl_b_w = doc.add_table(rows=17, cols=8)
    tbl_b_w.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_b_w.autofit = False
    
    headers_bw = ["Algorithm", "k", "Noise", "Silhouette", "DB Index", "ARI", "NMI", "V-Measure"]
    bw_widths = [Inches(1.8), Inches(0.5), Inches(0.6), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8)]
    
    for j in range(8):
        c = tbl_b_w.cell(0, j)
        c.width = bw_widths[j]
        set_cell_background(c, HEX_PRIMARY)
        set_cell_margins(c, top=80, bottom=80, left=60, right=60)
        p = c.paragraphs[0]
        r = p.add_run(headers_bw[j])
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)

    wisc_data = [
        ("15. Gaussian Mixture (GMM)", "2", "0.0%", "0.3145", "1.3770", "0.7740", "0.6611", "0.6611"),
        ("12. Spectral Clustering", "2", "0.0%", "0.3367", "1.3161", "0.7608", "0.6627", "0.6627"),
        ("16. Bayesian GMM (DPGMM)", "2", "0.0%", "0.3130", "1.3814", "0.7493", "0.6328", "0.6328"),
        ("2. Mini-Batch K-Means", "2", "0.0%", "0.3381", "1.3268", "0.6829", "0.5606", "0.5606"),
        ("1. K-Means", "2", "0.0%", "0.3434", "1.3205", "0.6536", "0.5324", "0.5324"),
        ("3. Bisecting K-Means", "2", "0.0%", "0.3434", "1.3205", "0.6536", "0.5324", "0.5324"),
        ("4. Agglomerative (Ward)", "2", "0.0%", "0.3394", "1.3700", "0.5750", "0.4569", "0.4569"),
        ("14. BIRCH", "2", "0.0%", "0.3394", "1.3700", "0.5750", "0.4569", "0.4569"),
        ("13. Affinity Propagation", "51", "0.0%", "0.0835", "1.3715", "0.0506", "0.2641", "0.2641"),
        ("6. Agglomerative (Average)", "2", "0.0%", "0.6340*", "0.6799*", "0.0073", "0.0151", "0.0151"),
        ("5. Agglomerative (Complete)", "2", "0.0%", "0.6607*", "0.4497*", "0.0048", "0.0102", "0.0102"),
        ("7. Agglomerative (Single)", "2", "0.0%", "0.6607*", "0.4497*", "0.0048", "0.0102", "0.0102"),
        ("8. DBSCAN", "1", "5.6%", "NaN", "NaN", "NaN", "NaN", "NaN"),
        ("9. HDBSCAN", "0", "100.0%", "NaN", "NaN", "NaN", "NaN", "NaN"),
        ("10. OPTICS", "1", "98.1%", "NaN", "NaN", "NaN", "NaN", "NaN"),
        ("11. Mean Shift", "1", "0.0%", "NaN", "NaN", "NaN", "NaN", "NaN")
    ]

    for row_idx, row in enumerate(wisc_data):
        bg = "FFFFFF"
        if row_idx == 0:
            bg = "EBF8FF" # Highlight top performer
        elif row_idx % 2 == 1:
            bg = HEX_LIGHT_BG
        for col_idx in range(8):
            cell = tbl_b_w.cell(row_idx + 1, col_idx)
            cell.width = bw_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            p = cell.paragraphs[0]
            if col_idx >= 1:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(row[col_idx])
            r.font.size = Pt(8)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = COLOR_PRIMARY
            elif col_idx == 5 and row_idx < 3: # Highlight top ARI
                r.font.bold = True
                r.font.color.rgb = RGBColor(43, 108, 176)
            else:
                r.font.color.rgb = COLOR_DARK

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(4)
    p_sp.paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # 2D PCA VISUALIZATION GRIDS & BENCHMARK CHARTS
    # -------------------------------------------------------------
    add_fig("plots/05_global_cancer_16_algorithms_pca.png", "Comparative 2D PCA Projections of Global Cancer Mortality Clustering across all 16 algorithms. K-Means, Ward, Spectral, and GMM generate balanced, coherent geopolitical quadrants, whereas Single/Complete linkage and Density algorithms collapse.", 5)

    add_fig("plots/06_wisconsin_16_algorithms_pca.png", "Comparative 2D PCA Projections of Wisconsin Breast Cancer Diagnostic Clustering across all 16 algorithms. Gaussian Mixture and Spectral Clustering delineate the non-linear malignant/benign interface with high fidelity.", 6)

    add_fig("plots/07_benchmark_global_cancer.png", "Clustering Benchmark Metrics on Global Cancer Mortality Profile (204 Nations). Comparison of Silhouette Score (left, higher is better) and Davies-Bouldin Index (right, lower is better).", 7)

    add_fig("plots/08_benchmark_wisconsin_accuracy.png", "Clustering Diagnostic Accuracy Benchmark on Breast Cancer Pathology (569 Tumors). Adjusted Rand Index (ARI, left) and Normalized Mutual Information (NMI, right) demonstrate clear superiority of Gaussian Mixture and Spectral Clustering.", 8)

    # -------------------------------------------------------------
    # 5. IN-DEPTH TECHNICAL DECONSTRUCTION
    # -------------------------------------------------------------
    add_h2("4.3 Deep-Dive Technical Deconstruction")
    
    add_h3("A. Why Gaussian Mixture Models & Spectral Clustering Excel in Diagnostic Pathology")
    add_p("In clinical diagnostic pathology, tumor morphology is inherently continuous and heterogeneous. Cell nuclei transitioning from normal tissue to high-grade malignant carcinoma undergo gradual alterations in chromatin texture, irregular nuclear contours, and enlargement. The biological data does not form rigid, equal-sized spherical clusters. Standard K-Means enforces spherical boundaries with isotropic variance, resulting in an ARI of 0.6536.")
    add_p("Gaussian Mixture Models (GMM) with full covariance matrices allow each cluster to possess an independent, unconstrained elliptical covariance tensor. This flexibility captures the elongated dispersion of malignant nuclei in multi-dimensional space, yielding an Adjusted Rand Index of 0.7740 (NMI = 0.6611). Furthermore, Spectral Clustering constructs a non-linear graph manifold via k-nearest neighbors (k=10), allowing it to follow the non-convex curvature of the cellular boundary and achieving an ARI of 0.7608.")

    add_h3("B. The Collapse of Density-Based Clustering (The Curse of Dimensionality)")
    add_p("A critical empirical finding is the total failure of density-based paradigms (DBSCAN, HDBSCAN, OPTICS) on the 30-dimensional Wisconsin biopsy dataset:")
    add_bullet("HDBSCAN labeled 100.0% of biopsy samples as noise (zero clusters found).")
    add_bullet("OPTICS classified 98.1% of samples as noise, isolating only a trivial sub-cluster.")
    add_bullet("DBSCAN identified only 1 cluster, failing to divide benign from malignant tumors and marking 5.6% as noise.")
    add_p("This phenomenon is a direct consequence of the 'Curse of Dimensionality'. In 30 dimensions, data space becomes exponentially sparse. The ratio of the distance to the nearest neighbor versus the farthest neighbor approaches 1 as dimensionality increases: lim_{d→∞} (dist_max - dist_min) / dist_min = 0. Consequently, local density differences vanish, preventing epsilon-neighborhood search algorithms from discerning density peaks and valleys.")

    add_h3("C. The Single-Linkage Chaining Vulnerability")
    add_p("Agglomerative clustering with single, complete, and average linkages recorded deceptive Silhouette scores (0.42 to 0.66) while failing completely on ground truth recovery (ARI < 0.008). Inspection of cluster assignments reveals that single linkage suffers from catastrophic chaining: 567 samples were assigned to Cluster 0, while samples #186 and #352 were partitioned into separate singletons. Because distance from singletons to the mega-cluster is massive, the mathematical formula for silhouette produces an artificially elevated score. This demonstrates that internal validation metrics must never be evaluated in isolation without inspecting cluster size parity.")

    # -------------------------------------------------------------
    # 6. EPIDEMIOLOGICAL CLUSTER PROFILES
    # -------------------------------------------------------------
    add_h1("5. Global Cancer Typology & Epidemiological Fingerprinting")

    add_p("To translate unsupervised clustering into actionable public health insights, the 204 sovereign nations were profiled using the balanced partitions generated by K-Means and Ward's hierarchical agglomeration at k = 4. Relative percentage deviation was calculated against the unweighted global baseline:")
    add_p("Relative Deviation (%) = [(Cluster Mean - Global Mean) / Global Mean] × 100%")

    add_fig("plots/09_cluster_fingerprint_heatmap.png", "Epidemiological Fingerprint Heatmap: Relative Cancer Mortality Deviation by Cluster (%) across 22 Cancer Types. Red cells indicate mortality rates significantly above the global average; blue cells indicate rates below the global average.", 9)

    add_h2("5.1 Quantitative Cluster Mortality Profiles")
    add_p("Table 4 details the absolute age-standardized mortality rates (deaths per 100,000) and percentage deviations from the global mean across all 22 oncological conditions:")

    # Table 4: Cluster Profiles
    tbl_prof = doc.add_table(rows=23, cols=5)
    tbl_prof.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_prof.autofit = False
    prof_widths = [Inches(2.0), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.2)]
    
    headers_prof = ["Cancer Type", "Cluster 0\n(Emerging, n=103)", "Cluster 1\n(Island/Afr, n=36)", "Cluster 2\n(Tonga, n=1)", "Cluster 3\n(Industrial, n=64)"]
    for j in range(5):
        c = tbl_prof.cell(0, j)
        c.width = prof_widths[j]
        set_cell_background(c, HEX_PRIMARY)
        set_cell_margins(c, top=80, bottom=80, left=60, right=60)
        p = c.paragraphs[0]
        r = p.add_run(headers_prof[j])
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)

    profiles_data = [
        ("Tracheal, bronchus, & lung", "14.91 (-27%)", "19.47 (-5%)", "29.48 (+44%)", "30.64 (+50%)"),
        ("Colon and rectum", "10.00 (-23%)", "14.17 (+9%)", "6.91 (-47%)", "19.08 (+47%)"),
        ("Stomach cancer", "11.35 (+7%)", "12.69 (+20%)", "13.85 (+31%)", "8.45 (-20%)"),
        ("Breast cancer", "9.32 (-20%)", "17.01 (+46%)", "17.46 (+50%)", "11.22 (-4%)"),
        ("Prostate cancer", "9.54 (-21%)", "21.89 (+81%)", "14.24 (+17%)", "9.56 (-21%)"),
        ("Pancreatic cancer", "4.65 (-31%)", "6.24 (-7%)", "4.49 (-33%)", "9.80 (+46%)"),
        ("Esophageal cancer", "5.08 (+6%)", "6.54 (+37%)", "2.21 (-54%)", "3.52 (-26%)"),
        ("Cervical cancer", "6.68 (+10%)", "11.24 (+85%)", "9.42 (+55%)", "2.38 (-61%)"),
        ("Liver cancer", "6.73 (-10%)", "6.35 (-15%)", "24.74 (+232%)", "4.37 (-41%)"),
        ("Brain & central nervous", "2.17 (-23%)", "1.98 (-30%)", "1.58 (-44%)", "4.58 (+62%)"),
        ("Bladder cancer", "2.53 (-21%)", "3.01 (-6%)", "1.79 (-44%)", "4.60 (+44%)"),
        ("Kidney cancer", "1.40 (-37%)", "1.61 (-27%)", "0.87 (-60%)", "3.71 (+67%)"),
        ("Ovarian cancer", "2.05 (-21%)", "3.06 (+18%)", "2.23 (-14%)", "3.54 (+36%)"),
        ("Gallbladder & biliary", "1.61 (-6%)", "1.56 (-9%)", "0.80 (-53%)", "1.93 (+13%)"),
        ("Larynx cancer", "1.32 (-12%)", "1.77 (+17%)", "0.65 (-57%)", "1.75 (+16%)"),
        ("Lip and oral cavity", "1.81 (-19%)", "3.74 (+67%)", "1.61 (-28%)", "1.96 (-12%)"),
        ("Uterine cancer", "1.24 (-26%)", "2.92 (+75%)", "1.92 (+15%)", "1.67 (+0%)"),
        ("Other pharynx", "0.67 (-23%)", "1.04 (+18%)", "0.27 (-69%)", "1.19 (+36%)"),
        ("Nasopharynx cancer", "0.76 (+10%)", "0.97 (+39%)", "0.93 (+34%)", "0.45 (-35%)"),
        ("Non-melanoma skin", "0.66 (-16%)", "0.81 (+3%)", "2.47 (+216%)", "0.83 (+6%)"),
        ("Thyroid cancer", "0.63 (+2%)", "0.84 (+35%)", "0.66 (+6%)", "0.56 (-10%)"),
        ("Testicular cancer", "0.12 (-44%)", "0.18 (-14%)", "6.87 (+3189%)", "0.22 (+5%)")
    ]

    for row_idx, row in enumerate(profiles_data):
        bg = HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
        for col_idx in range(5):
            cell = tbl_prof.cell(row_idx + 1, col_idx)
            cell.width = prof_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
            p = cell.paragraphs[0]
            if col_idx >= 1:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(row[col_idx])
            r.font.size = Pt(8)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = COLOR_PRIMARY
            else:
                r.font.color.rgb = COLOR_DARK

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(6)
    p_sp.paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # 5.2 EPIDEMIOLOGICAL INTERPRETATION & HEALTH POLICY
    # -------------------------------------------------------------
    add_h2("5.2 Epidemiological Typology Interpretation & Health Policy Actions")

    add_h3("Cluster 0: Emerging & Transitioning Economies (N = 103 Sovereign Nations)")
    add_p("Representative Nations: Brazil, India, Egypt, Algeria, Mexico, South Africa, Indonesia, Colombia, Bangladesh, Afghanistan.")
    add_p("Epidemiological Profile: Moderately elevated stomach cancer (11.35 deaths/100k, +7%) and esophageal cancer (5.08, +6%), with significantly depressed rates of lung (-27%), colorectal (-23%), pancreatic (-31%), and kidney cancer (-37%). This cluster reflects nations undergoing rapid economic transition where infectious/dietary risk factors (Helicobacter pylori, traditional salt preservation) persist, while western metabolic and tobacco-related chronic diseases have not yet reached peak mortality.")
    add_p("Health Policy Interventions: Mass population screening and eradication programs for Helicobacter pylori; refrigeration and food preservation standards; early endoscopic screening for gastric malignancies.")

    add_h3("Cluster 1: Sub-Saharan & Developing Island Nations (N = 36 Sovereign Nations)")
    add_p("Representative Nations: Botswana, Gabon, Fiji, Bahamas, Bolivia, Barbados, Cape Verde, Namibia, Papua New Guinea, Eswatini.")
    add_p("Epidemiological Profile: Massive elevations in reproductive and gynecological malignancies. Prostate cancer mortality is exceptionally high (21.89 deaths/100k, +81% above global baseline), cervical cancer is elevated by +85% (11.24 deaths/100k), and breast cancer mortality is elevated by +46% (17.01 deaths/100k). High mortality is driven not necessarily by higher biological incidence, but by delayed stage-at-diagnosis and lack of specialized radiotherapy, surgical oncology, and screening access.")
    add_p("Health Policy Interventions: Universal national HPV vaccination campaigns for adolescent girls; widespread visual inspection with acetic acid (VIA) and HPV DNA primary screening; regional decentralized prostate-specific antigen (PSA) diagnostic pathways; expansion of regional linear accelerators for radiotherapy.")

    add_h3("Cluster 2: Isolated Outlier Endemic Archetype (N = 1 Nation: Tonga)")
    add_p("Representative Nation: Kingdom of Tonga (Polynesian Island State).")
    add_p("Epidemiological Profile: Tonga forms a unique single-element cluster across K-Means, Mini-Batch K-Means, and Ward hierarchical grouping. It exhibits an extraordinary testicular cancer mortality rate of 6.87 deaths/100,000 (+3,189% above global average of 0.21), extreme liver cancer mortality (24.74 deaths/100k, +232%), non-melanoma skin cancer (2.47, +216%), and elevated lung cancer (29.48, +44%).")
    add_p("Health Policy Interventions: Targeted epidemiological field investigations to determine whether testicular and liver cancer spikes stem from rare regional genetic founder variants, endemic hepatitis B infection clusters, or severe localized environmental toxin/endocrine disruptor exposures; urgent deployment of ultrasound diagnostic units and platinum-based chemotherapy supply chains.")

    add_h3("Cluster 3: Industrialized & High-SDI Nations (N = 64 Sovereign Nations)")
    add_p("Representative Nations: Canada, Australia, Germany, United Kingdom, France, United States, Japan, Italy, Netherlands, South Korea.")
    add_p("Epidemiological Profile: Extreme elevation in chronic non-communicable malignancies associated with western diet, metabolic syndrome, high socio-demographic index (SDI), and historical tobacco consumption. Dominant drivers include lung cancer (30.64 deaths/100k, +50%), colorectal cancer (19.08, +47%), kidney cancer (3.71, +67%), brain/CNS malignancies (4.58, +62%), and pancreatic cancer (9.80, +46%). Conversely, cervical cancer is at an all-time low (2.38, -61% below baseline), reflecting decades of successful Pap smear cytology and HPV vaccination.")
    add_p("Health Policy Interventions: Low-dose computed tomography (LDCT) lung cancer screening for heavy smoking histories; nationwide colonoscopy and fecal immunochemical testing (FIT) starting at age 45; public health metabolic and ultra-processed food taxes; targeted smoking cessation.")

    # -------------------------------------------------------------
    # 7. CLINICAL TRANSLATION (WISCONSIN BENCHMARK)
    # -------------------------------------------------------------
    add_h1("6. Clinical Decision Support & Diagnostic Pathology Implications")

    add_p("The benchmark results on the Wisconsin Diagnostic Breast Cancer cohort provide critical insights for the deployment of unsupervised machine learning in computational pathology and clinical decision support:")

    add_h2("6.1 Unsupervised Malignancy Discovery")
    add_p("Without providing a single diagnostic label during training, Gaussian Mixture Models achieved an Adjusted Rand Index of 0.7740 and a Normalized Mutual Information of 0.6611. In clinical terms, GMM autonomously separated benign fibroadenomas from malignant invasive ductal carcinomas with over 90% raw concordance. This validates that multi-dimensional nuclear morphometry (nuclear area, perimeter, and concavity) contains an intrinsic geometric manifold that reflects true biological pathology.")

    add_h2("6.2 Probabilistic Soft Assignments vs Hard Boundaries")
    add_p("A fundamental advantage of GMM over centroid methods (K-Means) is its computation of continuous posterior probabilities for each tissue sample:")
    add_p("P(Class = Malignant | x) = [π_m · N(x | μ_m, Σ_m)] / ∑_k [π_k · N(x | μ_k, Σ_k)]")
    add_p("In clinical oncology, diagnostic ambiguity is common in borderline atypia, ductal carcinoma in situ (DCIS) with micro-invasion, and complex sclerosing lesions. Hard clustering algorithms force a binary categorization, concealing diagnostic uncertainty. In contrast, GMM flags samples where P(Malignant | x) falls between 0.40 and 0.60 as borderline cases, automatically directing them to senior human histopathologists for secondary immunohistochemical (IHC) or fluorescence in situ hybridization (FISH) staining.")

    # -------------------------------------------------------------
    # 8. ENGINEERING & DEPLOYMENT ARCHITECTURE
    # -------------------------------------------------------------
    add_h1("7. Engineering Architecture & Algorithmic Selection Guide")

    add_p("For machine learning engineers, clinical bioinformaticians, and healthcare data architects, selecting the appropriate clustering algorithm requires balancing cluster geometric assumptions, dimensionality, sample size, and latency. The following decision matrix synthesizes empirical findings into concrete deployment rules:")

    add_h2("7.1 Algorithmic Selection Decision Framework")
    add_bullet("Scenario 1: High-Dimensional Continuous Diagnostic Data (e.g., Cellular Pathology, Omics, Flow Cytometry) → RECOMMENDATION: Gaussian Mixture Models (GMM) or Spectral Clustering. Avoid DBSCAN/HDBSCAN entirely due to density collapse in spaces with >15 dimensions.")
    add_bullet("Scenario 2: Large-Scale Population or Patient Registry Data (N > 100,000, D < 15) → RECOMMENDATION: Mini-Batch K-Means or BIRCH. Both provide sub-linear or linear O(n) computational complexity, low memory overhead, and rapid convergence.")
    add_bullet("Scenario 3: Exploratory Taxonomy & Geopolitical Stratification (N < 2,000, Multi-Factorial Continuous) → RECOMMENDATION: Hierarchical Agglomerative Clustering with Ward Linkage. Produces an interpretable dendrogram and compact, balanced clusters without chaining artifacts.")
    add_bullet("Scenario 4: Spatial Epidemiology & Infectious Disease Outbreak Detection (2D/3D Geographic Coordinates) → RECOMMENDATION: HDBSCAN. In low-dimensional Euclidean spaces (latitude, longitude, time), HDBSCAN excels at discovering irregularly shaped geographical outbreak clusters without prior knowledge of k.")

    # -------------------------------------------------------------
    # 9. LIMITATIONS & FUTURE DIRECTIONS
    # -------------------------------------------------------------
    add_h1("8. Methodological Limitations & Future Research Directions")

    add_p("While this study presents an exhaustive benchmark across 16 algorithmic architectures, several methodological boundaries must be acknowledged:")
    add_bullet("Cross-Sectional Observational Nature: The global epidemiological dataset reflects cross-sectional mortality in 2019. It does not capture longitudinal trajectories, historical policy shifts, or post-2020 pandemic disruptions to oncology screening.")
    add_bullet("Under-Reporting and Diagnostic Disparities: In lower-middle-income nations (LMICs), cancer registries often suffer from under-ascertainment, missing rural data, and diagnostic coding discrepancies, which may suppress recorded mortality rates.")
    add_bullet("Absence of Multi-Omics Features: The cellular pathology benchmark was restricted to nuclear morphometry. Future investigations will incorporate spatial transcriptomics, single-cell RNA sequencing (scRNA-seq), and genomic mutational signatures (e.g., BRCA1/2, TP53).")
    add_bullet("Deep Unsupervised Clustering: Promising future research includes evaluating Deep Embedded Clustering (DEC) and Variational Autoencoder (VAE) latent space clustering to learn non-linear manifold embeddings prior to mixture modeling.")

    # -------------------------------------------------------------
    # 10. REFERENCES & APPENDIX
    # -------------------------------------------------------------
    add_h1("9. References & Bibliographic Citations")

    refs = [
        "Ankerst, M., Breunig, M. M., Kriegel, H. P., & Sander, J. (1999). OPTICS: Ordering points to identify the clustering structure. ACM SIGMOD Record, 28(2), 49-60.",
        "Campello, R. J., Moulavi, D., & Sander, J. (2013). Density-based clustering based on hierarchical density estimates. In Pacific-Asia Conference on Knowledge Discovery and Data Mining (pp. 160-172). Springer, Berlin, Heidelberg.",
        "Comaniciu, D., & Meer, P. (2002). Mean shift: A robust approach toward feature space analysis. IEEE Transactions on Pattern Analysis and Machine Intelligence, 24(5), 603-619.",
        "Dempster, A. P., Laird, N. M., & Rubin, D. B. (1977). Maximum likelihood from incomplete data via the EM algorithm. Journal of the Royal Statistical Society: Series B, 39(1), 1-22.",
        "Ester, M., Kriegel, H. P., Sander, J., & Xu, X. (1996). A density-based algorithm for discovering clusters in large spatial databases with noise. In KDD (Vol. 96, No. 34, pp. 226-233).",
        "Frey, B. J., & Dueck, D. (2007). Clustering by passing messages between data points. Science, 315(5814), 972-976.",
        "Institute for Health Metrics and Evaluation (IHME). Global Burden of Disease Study 2019 (GBD 2019) Results. Seattle, United States: IHME, 2020.",
        "Lloyd, S. (1982). Least squares quantization in PCM. IEEE Transactions on Information Theory, 28(2), 129-137.",
        "MacQueen, J. (1967). Some methods for classification and analysis of multivariate observations. In Proceedings of the Fifth Berkeley Symposium on Mathematical Statistics and Probability (Vol. 1, No. 14, pp. 281-297).",
        "Ng, A. Y., Jordan, M. I., & Weiss, Y. (2001). On spectral clustering: Analysis and an algorithm. Advances in Neural Information Processing Systems, 14, 849-856.",
        "Rousseeuw, P. J. (1987). Silhouettes: A graphical aid to the interpretation and validation of cluster analysis. Journal of Computational and Applied Mathematics, 20, 53-65.",
        "Sculley, D. (2010). Web-scale k-means clustering. In Proceedings of the 19th International Conference on World Wide Web (pp. 1177-1178).",
        "Street, W. N., Wolberg, W. H., & Mangasarian, O. L. (1993). Nuclear feature extraction for breast tumor diagnosis. In Biomedical Image Processing and Biomedical Visualization (Vol. 1905, pp. 861-870). SPIE.",
        "Ward Jr, J. H. (1963). Hierarchical grouping to optimize an objective function. Journal of the American Statistical Association, 58(301), 236-244.",
        "Zhang, T., Ramakrishnan, R., & Livny, M. (1996). BIRCH: An efficient data clustering method for very large databases. ACM SIGMOD Record, 25(2), 103-114."
    ]

    for ref in refs:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.left_indent = Inches(0.3)
        p_r.paragraph_format.first_line_indent = Inches(-0.3)
        p_r.paragraph_format.space_before = Pt(1)
        p_r.paragraph_format.space_after = Pt(3)
        r = p_r.add_run(ref)
        r.font.size = Pt(8.5)
        r.font.color.rgb = COLOR_DARK

    output_path = "/home/sashank/USLA/Cancer_Clustering_Analysis_Report.docx"
    print(f"Saving final report to {output_path}...")
    doc.save(output_path)
    print("Report generated successfully!")

if __name__ == '__main__':
    build_docx_report()
