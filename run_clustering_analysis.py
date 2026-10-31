import os
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.cluster.hierarchy import dendrogram, linkage
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.manifold import TSNE
from sklearn.neighbors import NearestNeighbors

from sklearn.cluster import (
    KMeans, MiniBatchKMeans, BisectingKMeans,
    AgglomerativeClustering, DBSCAN, HDBSCAN, OPTICS,
    MeanShift, estimate_bandwidth, SpectralClustering,
    AffinityPropagation, Birch
)
from sklearn.mixture import GaussianMixture, BayesianGaussianMixture
from sklearn.metrics import (
    silhouette_score, davies_bouldin_score, calinski_harabasz_score,
    adjusted_rand_score, normalized_mutual_info_score,
    homogeneity_score, completeness_score, v_measure_score
)
from sklearn.datasets import load_breast_cancer

# Set plotting styles
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['figure.autolayout'] = True

os.makedirs('plots', exist_ok=True)
os.makedirs('clustering_results', exist_ok=True)

print('=== 1. PREPARING DATASETS ===')

# --- Dataset 1: Global Cancer Mortality by Type (Macro Population Level) ---
df_global_raw = pd.read_csv('archive/cancer-death-rates-by-type.csv')
# Filter for year 2019 and sovereign countries with valid 3-letter ISO code
df_global = df_global_raw[(df_global_raw['Year'] == 2019) & (df_global_raw['Code'].notnull()) & (df_global_raw['Code'].str.len() == 3)].copy()
# Reset index
df_global = df_global.reset_index(drop=True)

feature_cols = [c for c in df_global.columns if c not in ['Entity', 'Code', 'Year']]
# Clean column names for readability
rename_dict = {}
for col in feature_cols:
    short = col.replace('Deaths - ', '').replace(' - Sex: Both - Age: Age-standardized (Rate)', '').strip()
    rename_dict[col] = short
df_global = df_global.rename(columns=rename_dict)
clean_features = list(rename_dict.values())

print(f'Global Cancer Dataset: {df_global.shape[0]} countries, {len(clean_features)} cancer mortality rates')

X_global = df_global[clean_features].values
scaler_global = StandardScaler()
X_global_scaled = scaler_global.fit_transform(X_global)

# PCA for Global Dataset
pca_global = PCA(n_components=2, random_state=42)
X_global_pca = pca_global.fit_transform(X_global_scaled)
print(f'Global Dataset PCA explained variance ratio (2 components): {pca_global.explained_variance_ratio_} (Total: {pca_global.explained_variance_ratio_.sum():.2%})')

# --- Dataset 2: Wisconsin Diagnostic Breast Cancer (Micro Cellular Level) ---
cancer_wisc = load_breast_cancer()
X_wisc = cancer_wisc.data
y_wisc = cancer_wisc.target # 0: Malignant, 1: Benign
wisc_feature_names = cancer_wisc.feature_names
wisc_target_names = cancer_wisc.target_names

print(f'Wisconsin Breast Cancer Dataset: {X_wisc.shape[0]} samples, {X_wisc.shape[1]} cellular features')
print(f'Target distribution: Malignant (0): {(y_wisc == 0).sum()}, Benign (1): {(y_wisc == 1).sum()}')

scaler_wisc = StandardScaler()
X_wisc_scaled = scaler_wisc.fit_transform(X_wisc)

pca_wisc = PCA(n_components=2, random_state=42)
X_wisc_pca = pca_wisc.fit_transform(X_wisc_scaled)
print(f'Wisconsin PCA explained variance ratio (2 components): {pca_wisc.explained_variance_ratio_} (Total: {pca_wisc.explained_variance_ratio_.sum():.2%})')

print('\n=== 2. EXPLORATORY ANALYSIS & OPTIMAL K DETERMINATION ===')

# Determine optimal k for global dataset using Elbow & Silhouette
k_range = range(2, 11)
inertias_global = []
silhouettes_global = []
db_scores_global = []

for k in k_range:
    km = KMeans(n_clusters=k, random_state=42, n_init=10)
    labels = km.fit_predict(X_global_scaled)
    inertias_global.append(km.inertia_)
    silhouettes_global.append(silhouette_score(X_global_scaled, labels))
    db_scores_global.append(davies_bouldin_score(X_global_scaled, labels))

optimal_k_global = k_range[np.argmax(silhouettes_global)]
print(f'Optimal k for Global Dataset by Silhouette: k={optimal_k_global} (Score: {max(silhouettes_global):.4f})')

fig, ax1 = plt.subplots(1, 2, figsize=(14, 5))
ax1[0].plot(k_range, inertias_global, 'bo-', linewidth=2, markersize=8)
ax1[0].set_title('Elbow Method (Inertia vs k) - Global Cancer Data', fontsize=12, fontweight='bold')
ax1[0].set_xlabel('Number of Clusters (k)', fontsize=11)
ax1[0].set_ylabel('Inertia (Sum of Squared Distances)', fontsize=11)

ax1[1].plot(k_range, silhouettes_global, 'gs-', linewidth=2, markersize=8, label='Silhouette Score')
ax1[1].set_title('Silhouette Score vs k - Global Cancer Data', fontsize=12, fontweight='bold')
ax1[1].set_xlabel('Number of Clusters (k)', fontsize=11)
ax1[1].set_ylabel('Silhouette Score', fontsize=11)
plt.tight_layout()
plt.savefig('plots/01_optimal_k_global_cancer.png', dpi=300)
plt.close()

# For Wisconsin dataset: natural k=2 (malignant vs benign), but let's test k=2 to 10
inertias_wisc = []
silhouettes_wisc = []
for k in k_range:
    km = KMeans(n_clusters=k, random_state=42, n_init=10)
    labels = km.fit_predict(X_wisc_scaled)
    inertias_wisc.append(km.inertia_)
    silhouettes_wisc.append(silhouette_score(X_wisc_scaled, labels))

fig, ax2 = plt.subplots(1, 2, figsize=(14, 5))
ax2[0].plot(k_range, inertias_wisc, 'ro-', linewidth=2, markersize=8)
ax2[0].set_title('Elbow Method - Wisconsin Breast Cancer', fontsize=12, fontweight='bold')
ax2[0].set_xlabel('k', fontsize=11)
ax2[0].set_ylabel('Inertia', fontsize=11)

ax2[1].plot(k_range, silhouettes_wisc, 'ms-', linewidth=2, markersize=8)
ax2[1].set_title('Silhouette Score vs k - Wisconsin Breast Cancer', fontsize=12, fontweight='bold')
ax2[1].set_xlabel('k', fontsize=11)
ax2[1].set_ylabel('Silhouette Score', fontsize=11)
plt.tight_layout()
plt.savefig('plots/02_optimal_k_wisconsin.png', dpi=300)
plt.close()

# Dendrogram for Global Cancer Dataset
plt.figure(figsize=(16, 7))
linked = linkage(X_global_scaled, method='ward')
dendrogram(linked, orientation='top', labels=df_global['Code'].values, distance_sort='descending', show_leaf_counts=True, leaf_font_size=6)
plt.title('Hierarchical Clustering Dendrogram (Ward Linkage) - Global Cancer Mortality', fontsize=14, fontweight='bold')
plt.xlabel('Country Code', fontsize=11)
plt.ylabel('Ward Distance', fontsize=11)
plt.tight_layout()
plt.savefig('plots/03_dendrogram_global_cancer.png', dpi=300)
plt.close()

# Correlation Heatmap among Cancer Types
plt.figure(figsize=(14, 11))
corr = df_global[clean_features].corr()
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, cmap='coolwarm', vmin=-1, vmax=1, annot=True, fmt='.2f', annot_kws={'size': 7}, cbar_kws={'shrink': 0.8})
plt.title('Global Correlation Matrix of Cancer Mortality Types (2019)', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('plots/04_cancer_types_correlation_matrix.png', dpi=300)
plt.close()

print('\n=== 3. DEFINING AND EXECUTING ALL 16 CLUSTERING ALGORITHMS ===')

def evaluate_clustering(X, labels, y_true=None):
    """Computes internal and external clustering metrics."""
    # Filter noise points (label == -1) for internal metrics if necessary
    valid_mask = (labels != -1)
    n_clusters = len(set(labels)) - (1 if -1 in labels else 0)
    noise_ratio = (labels == -1).sum() / len(labels) if -1 in labels else 0.0

    if n_clusters < 2:
        return {
            'n_clusters': n_clusters,
            'noise_ratio': noise_ratio,
            'silhouette': np.nan,
            'davies_bouldin': np.nan,
            'calinski_harabasz': np.nan,
            'ari': np.nan,
            'nmi': np.nan,
            'v_measure': np.nan
        }

    # Internal metrics on valid points if noise present, or on full data
    X_eval = X[valid_mask] if noise_ratio > 0 else X
    labels_eval = labels[valid_mask] if noise_ratio > 0 else labels

    if len(set(labels_eval)) < 2:
        sil = np.nan
        db = np.nan
        ch = np.nan
    else:
        sil = silhouette_score(X_eval, labels_eval)
        db = davies_bouldin_score(X_eval, labels_eval)
        ch = calinski_harabasz_score(X_eval, labels_eval)

    res = {
        'n_clusters': n_clusters,
        'noise_ratio': noise_ratio,
        'silhouette': sil,
        'davies_bouldin': db,
        'calinski_harabasz': ch
    }

    if y_true is not None:
        res['ari'] = adjusted_rand_score(y_true, labels)
        res['nmi'] = normalized_mutual_info_score(y_true, labels)
        res['v_measure'] = v_measure_score(y_true, labels)

    return res

def run_all_algorithms(X_scaled, target_k=4, y_true=None, dataset_name='global'):
    # Determine density parameters tuned for dataset
    if dataset_name == 'global':
        # 204 points, 22 dims
        eps_db = 3.6
        min_samp_db = 4
        min_cluster_hdb = 5
        bw = estimate_bandwidth(X_scaled, quantile=0.25)
    else:
        # 569 points, 30 dims
        eps_db = 4.5
        min_samp_db = 5
        min_cluster_hdb = 15
        bw = estimate_bandwidth(X_scaled, quantile=0.20)

    algorithms = {
        '1. K-Means': KMeans(n_clusters=target_k, random_state=42, n_init=10),
        '2. Mini-Batch K-Means': MiniBatchKMeans(n_clusters=target_k, random_state=42, n_init=10, batch_size=64),
        '3. Bisecting K-Means': BisectingKMeans(n_clusters=target_k, random_state=42, n_init=10),
        '4. Agglomerative (Ward)': AgglomerativeClustering(n_clusters=target_k, linkage='ward'),
        '5. Agglomerative (Complete)': AgglomerativeClustering(n_clusters=target_k, linkage='complete'),
        '6. Agglomerative (Average)': AgglomerativeClustering(n_clusters=target_k, linkage='average'),
        '7. Agglomerative (Single)': AgglomerativeClustering(n_clusters=target_k, linkage='single'),
        '8. DBSCAN': DBSCAN(eps=eps_db, min_samples=min_samp_db),
        '9. HDBSCAN': HDBSCAN(min_cluster_size=min_cluster_hdb),
        '10. OPTICS': OPTICS(min_samples=min_samp_db, xi=0.05),
        '11. Mean Shift': MeanShift(bandwidth=bw, bin_seeding=True),
        '12. Spectral Clustering': SpectralClustering(n_clusters=target_k, random_state=42, affinity='nearest_neighbors', n_neighbors=10),
        '13. Affinity Propagation': AffinityPropagation(random_state=42, damping=0.85, max_iter=400),
        '14. BIRCH': Birch(n_clusters=target_k, threshold=0.5),
        '15. Gaussian Mixture (GMM)': GaussianMixture(n_components=target_k, random_state=42, covariance_type='full'),
        '16. Bayesian GMM (DPGMM)': BayesianGaussianMixture(n_components=target_k, random_state=42, covariance_type='full', weight_concentration_prior_type='dirichlet_process')
    }

    results = []
    labels_dict = {}

    for name, model in algorithms.items():
        try:
            if hasattr(model, 'fit_predict'):
                labels = model.fit_predict(X_scaled)
            elif hasattr(model, 'predict'):
                model.fit(X_scaled)
                labels = model.predict(X_scaled)
            else:
                model.fit(X_scaled)
                labels = model.labels_

            metrics = evaluate_clustering(X_scaled, labels, y_true=y_true)
            metrics['Algorithm'] = name
            results.append(metrics)
            labels_dict[name] = labels
            print(f"[{dataset_name.upper()}] {name:30s} -> Clusters: {metrics['n_clusters']:2d} | Sil: {metrics['silhouette']:.4f} | DB: {metrics['davies_bouldin']:.4f} | Noise: {metrics['noise_ratio']:.1%}")
        except Exception as e:
            print(f"[{dataset_name.upper()}] Error running {name}: {e}")
            metrics = {
                'Algorithm': name,
                'n_clusters': np.nan,
                'noise_ratio': np.nan,
                'silhouette': np.nan,
                'davies_bouldin': np.nan,
                'calinski_harabasz': np.nan
            }
            if y_true is not None:
                metrics['ari'] = np.nan
                metrics['nmi'] = np.nan
                metrics['v_measure'] = np.nan
            results.append(metrics)

    df_res = pd.DataFrame(results)
    # Reorder columns
    cols = ['Algorithm', 'n_clusters', 'noise_ratio', 'silhouette', 'davies_bouldin', 'calinski_harabasz']
    if y_true is not None:
        cols.extend(['ari', 'nmi', 'v_measure'])
    df_res = df_res[cols]
    return df_res, labels_dict

print('\n--- Running 16 Algorithms on Global Cancer Dataset (k=4) ---')
# k=4 captures major epidemiological global typologies
results_global, labels_global = run_all_algorithms(X_global_scaled, target_k=4, dataset_name='global')
results_global.to_csv('clustering_results/global_cancer_clustering_comparison.csv', index=False)

print('\n--- Running 16 Algorithms on Wisconsin Diagnostic Breast Cancer (k=2) ---')
results_wisc, labels_wisc = run_all_algorithms(X_wisc_scaled, target_k=2, y_true=y_wisc, dataset_name='wisconsin')
results_wisc.to_csv('clustering_results/wisconsin_breast_cancer_clustering_comparison.csv', index=False)

print('\n=== 4. PLOTTING CLUSTER VISUALIZATIONS ===')

# Plot 16 algorithms PCA grid for Global Cancer
fig, axes = plt.subplots(4, 4, figsize=(20, 18))
axes = axes.flatten()

for idx, (name, labels) in enumerate(labels_global.items()):
    ax = axes[idx]
    # Unique labels
    unique_labels = sorted(list(set(labels)))
    colors = plt.cm.tab10(np.linspace(0, 1, max(len(unique_labels), 10)))
    
    for l_idx, lbl in enumerate(unique_labels):
        mask = (labels == lbl)
        if lbl == -1:
            ax.scatter(X_global_pca[mask, 0], X_global_pca[mask, 1], c='lightgray', marker='x', s=25, alpha=0.6, label='Noise')
        else:
            color = colors[l_idx % 10]
            ax.scatter(X_global_pca[mask, 0], X_global_pca[mask, 1], c=[color], marker='o', s=35, alpha=0.8, edgecolors='none', label=f'C{lbl}')
            
    ax.set_title(name, fontsize=10, fontweight='bold')
    ax.set_xticks([])
    ax.set_yticks([])

plt.suptitle('Global Cancer Mortality Clustering (PCA 2D Projection across 16 Algorithms)', fontsize=16, fontweight='bold')
plt.tight_layout(rect=[0, 0.03, 1, 0.96])
plt.savefig('plots/05_global_cancer_16_algorithms_pca.png', dpi=300)
plt.close()

# Plot 16 algorithms PCA grid for Wisconsin Breast Cancer
fig, axes = plt.subplots(4, 4, figsize=(20, 18))
axes = axes.flatten()

for idx, (name, labels) in enumerate(labels_wisc.items()):
    ax = axes[idx]
    unique_labels = sorted(list(set(labels)))
    colors = plt.cm.Set1(np.linspace(0, 1, max(len(unique_labels), 9)))
    
    for l_idx, lbl in enumerate(unique_labels):
        mask = (labels == lbl)
        if lbl == -1:
            ax.scatter(X_wisc_pca[mask, 0], X_wisc_pca[mask, 1], c='lightgray', marker='x', s=20, alpha=0.5, label='Noise')
        else:
            color = colors[l_idx % 9]
            ax.scatter(X_wisc_pca[mask, 0], X_wisc_pca[mask, 1], c=[color], marker='o', s=25, alpha=0.7, edgecolors='none', label=f'C{lbl}')
            
    ax.set_title(name, fontsize=10, fontweight='bold')
    ax.set_xticks([])
    ax.set_yticks([])

plt.suptitle('Wisconsin Breast Cancer Cellular Clustering (PCA 2D across 16 Algorithms)', fontsize=16, fontweight='bold')
plt.tight_layout(rect=[0, 0.03, 1, 0.96])
plt.savefig('plots/06_wisconsin_16_algorithms_pca.png', dpi=300)
plt.close()

# Plot Performance Benchmark Bars for Global Cancer (Silhouette & Davies-Bouldin)
fig, ax = plt.subplots(1, 2, figsize=(16, 6))
df_plot_global = results_global.dropna(subset=['silhouette']).sort_values('silhouette', ascending=False)
sns.barplot(data=df_plot_global, x='silhouette', y='Algorithm', palette='viridis', ax=ax[0])
ax[0].set_title('Silhouette Score (Higher is Better)', fontsize=12, fontweight='bold')
ax[0].set_xlabel('Silhouette Score')
ax[0].set_ylabel('')

df_plot_db = results_global.dropna(subset=['davies_bouldin']).sort_values('davies_bouldin', ascending=True)
sns.barplot(data=df_plot_db, x='davies_bouldin', y='Algorithm', palette='mako', ax=ax[1])
ax[1].set_title('Davies-Bouldin Index (Lower is Better)', fontsize=12, fontweight='bold')
ax[1].set_xlabel('Davies-Bouldin Index')
ax[1].set_ylabel('')

plt.suptitle('Clustering Benchmark on Global Cancer Mortality Profile (204 Nations)', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('plots/07_benchmark_global_cancer.png', dpi=300)
plt.close()

# Plot Performance Benchmark Bars for Wisconsin (ARI & NMI vs Ground Truth)
fig, ax = plt.subplots(1, 2, figsize=(16, 6))
df_plot_ari = results_wisc.dropna(subset=['ari']).sort_values('ari', ascending=False)
sns.barplot(data=df_plot_ari, x='ari', y='Algorithm', palette='magma', ax=ax[0])
ax[0].set_title('Adjusted Rand Index (ARI vs Diagnosis) - Higher is Better', fontsize=12, fontweight='bold')
ax[0].set_xlabel('Adjusted Rand Index')
ax[0].set_ylabel('')

df_plot_nmi = results_wisc.dropna(subset=['nmi']).sort_values('nmi', ascending=False)
sns.barplot(data=df_plot_nmi, x='nmi', y='Algorithm', palette='rocket', ax=ax[1])
ax[1].set_title('Normalized Mutual Information (NMI) - Higher is Better', fontsize=12, fontweight='bold')
ax[1].set_xlabel('Normalized Mutual Information')
ax[1].set_ylabel('')

plt.suptitle('Clustering Diagnostic Accuracy Benchmark on Breast Cancer Pathology (569 Tumors)', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('plots/08_benchmark_wisconsin_accuracy.png', dpi=300)
plt.close()

print('\n=== 5. DETAILED EPIDEMIOLOGICAL CLUSTER PROFILING ===')
# Profile the 4 clusters identified by K-Means / Ward on Global Data
df_global['Cluster_KMeans'] = labels_global['1. K-Means']
df_global['Cluster_Ward'] = labels_global['4. Agglomerative (Ward)']

# Calculate mean mortality rate for each cancer type per cluster
cluster_profile = df_global.groupby('Cluster_KMeans')[clean_features].mean()
cluster_profile.to_csv('clustering_results/global_cancer_kmeans_cluster_profiles.csv')

# Top 5 distinctive cancer types per cluster compared to global average
global_means = df_global[clean_features].mean()
relative_diff = (cluster_profile - global_means) / global_means

print('Distinctive Cancer Characteristics by Cluster:')
for c_id in range(4):
    top_elevated = relative_diff.loc[c_id].sort_values(ascending=False).head(4)
    top_reduced = relative_diff.loc[c_id].sort_values().head(3)
    sample_countries = df_global[df_global['Cluster_KMeans'] == c_id]['Entity'].head(8).tolist()
    print(f"\n--- Cluster {c_id} (n={(df_global['Cluster_KMeans'] == c_id).sum()} countries) ---")
    print(f"Sample Countries: {', '.join(sample_countries)}")
    print("Highest Elevated Mortality Rates (% above global average):")
    for cancer, val in top_elevated.items():
        print(f"  * {cancer}: +{val*100:.1f}% ({cluster_profile.loc[c_id, cancer]:.2f} vs global avg {global_means[cancer]:.2f})")
    print("Lowest Mortality Rates (% below global average):")
    for cancer, val in top_reduced.items():
        print(f"  * {cancer}: {val*100:.1f}% ({cluster_profile.loc[c_id, cancer]:.2f} vs global avg {global_means[cancer]:.2f})")

# Heatmap of relative differences by cluster
plt.figure(figsize=(14, 6))
sns.heatmap(relative_diff.T * 100, cmap='vlag', center=0, annot=True, fmt='.1f', cbar_kws={'label': '% Difference from Global Mean'})
plt.title('Epidemiological Fingerprint: Relative Cancer Mortality Deviation by Cluster (%)', fontsize=14, fontweight='bold')
plt.xlabel('Cluster ID', fontsize=12)
plt.ylabel('Cancer Type', fontsize=12)
plt.tight_layout()
plt.savefig('plots/09_cluster_fingerprint_heatmap.png', dpi=300)
plt.close()

# Save Country cluster assignments
country_clusters = df_global[['Entity', 'Code', 'Cluster_KMeans', 'Cluster_Ward']]
country_clusters.to_csv('clustering_results/country_cluster_assignments.csv', index=False)

print('\nAnalysis completed successfully!')
