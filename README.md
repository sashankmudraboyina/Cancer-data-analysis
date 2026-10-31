# Comparative Clustering Analysis and Algorithmic Benchmarking in Oncology

A dual-scale unsupervised machine learning investigation into global epidemiological cancer mortality profiles (**204 sovereign nations**, 22 cancer types) and microscopic cellular pathology (**569 fine-needle aspirate biopsies**, 30 continuous morphometric features) across **16 clustering architectures** and **6 algorithmic families**.

---

## 📄 Comprehensive Report

The complete, publication-grade research report is available in `.docx` format:
- **[`Cancer_Clustering_Analysis_Report.docx`](Cancer_Clustering_Analysis_Report.docx)**

---

## 🔬 Project Overview

This study benchmarks 16 clustering algorithms across two complementary observational scales in oncology:

1. **Macro-Epidemiological Scale**: 204 sovereign nations evaluated across 22 age-standardized cancer mortality rates per 100,000 population (IHME / Global Burden of Disease study, 2019 cross-section).
2. **Micro-Pathological Scale**: 569 fine-needle aspirates from the Wisconsin Diagnostic Breast Cancer (WDBC) cohort, parameterized into 30 continuous nuclear morphological descriptors with ground truth labels (357 Benign vs 212 Malignant).

---

## 🧠 Evaluated Algorithms (16 Models across 6 Families)

| Family | Algorithms |
| :--- | :--- |
| **Centroid-Based** | K-Means, Mini-Batch K-Means, Bisecting K-Means |
| **Hierarchical Agglomerative** | Ward Linkage, Complete Linkage, Average Linkage, Single Linkage |
| **Density-Based** | DBSCAN, HDBSCAN, OPTICS |
| **Mode-Seeking / Distribution Shift** | Mean Shift |
| **Graph & Manifold** | Spectral Clustering |
| **Exemplar & Tree-Subspace** | Affinity Propagation, BIRCH |
| **Probabilistic & Bayesian** | Gaussian Mixture Models (GMM), Bayesian GMM (DPGMM) |

---

## 📊 Key Findings

- **Top Cellular Diagnostic Performance**: **Gaussian Mixture Models (GMM)** achieved the highest unsupervised concordance with histological ground truth (**ARI = 0.7740**, **NMI = 0.6611**), followed closely by **Spectral Clustering** (**ARI = 0.7608**).
- **The Curse of Dimensionality in Clinical Density Clustering**: Density-based methods (DBSCAN, HDBSCAN, OPTICS) collapsed on 30-dimensional continuous cellular data due to uniform spatial density dispersion, marking between 5.6% and 100% of samples as unclassifiable noise.
- **Single-Linkage Chaining Artifact**: Agglomerative single linkage achieved an artificially inflated Silhouette score (0.6607) due to catastrophic chaining (1 mega-cluster + 2 singletons; ARI = 0.0048).
- **Four Global Epidemiological Typologies**:
  - **Cluster 0 (103 countries - Emerging Economies)**: Elevated stomach (+7%) and esophageal (+6%) mortality; lower lifestyle cancers.
  - **Cluster 1 (36 countries - Sub-Saharan & Island Nations)**: Severe gynecological and reproductive mortality burden (Prostate +81%, Cervical +85%, Breast +46%).
  - **Cluster 2 (1 country - Tonga / Geographic Outlier)**: Massive spikes in Testicular (+3,189%) and Liver (+232%) mortality.
  - **Cluster 3 (64 countries - Industrialized High-SDI)**: Severe metabolic/lifestyle burden (Lung +50%, Colorectal +47%, Pancreatic +46%, Kidney +67%, Brain +62%).

---

## 📁 Repository Structure

```text
├── Cancer_Clustering_Analysis_Report.docx  # Full formatted research monograph (.docx)
├── run_clustering_analysis.py              # Main clustering execution & benchmark script
├── create_report.py                        # Automated publication report generator
├── clustering_results/                     # Benchmark CSVs and cluster profiles
│   ├── country_cluster_assignments.csv
│   ├── global_cancer_clustering_comparison.csv
│   ├── global_cancer_kmeans_cluster_profiles.csv
│   └── wisconsin_breast_cancer_clustering_comparison.csv
├── plots/                                  # High-resolution analysis figures
│   ├── 01_optimal_k_global_cancer.png
│   ├── 02_optimal_k_wisconsin.png
│   ├── 03_dendrogram_global_cancer.png
│   ├── 04_cancer_types_correlation_matrix.png
│   ├── 05_global_cancer_16_algorithms_pca.png
│   ├── 06_wisconsin_16_algorithms_pca.png
│   ├── 07_benchmark_global_cancer.png
│   ├── 08_benchmark_wisconsin_accuracy.png
│   └── 09_cluster_fingerprint_heatmap.png
├── archive/                                # Underlying raw epidemiology datasets
└── README.md
```

---

## 🚀 Getting Started

### 1. Requirements

```bash
pip install numpy pandas scikit-learn scipy matplotlib seaborn python-docx
```

### 2. Run the Analysis

```bash
python run_clustering_analysis.py
```

### 3. Generate the Report

```bash
python create_report.py
```
