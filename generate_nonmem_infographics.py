#!/usr/bin/env python3
"""
NONMEM IRT Model Infographics Generator
Generates professional pharmacometric visualizations from NONMEM output file

Model: Cladribine in Multiple Sclerosis - IRT Model for EDSS Subscores
Reference: Novakovic et al., 2016
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, FancyArrowPatch
from matplotlib.gridspec import GridSpec
import numpy as np
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap
import warnings
warnings.filterwarnings('ignore')

# Set style for professional appearance
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['figure.facecolor'] = 'white'

# ============================================================================
# PARAMETER VALUES FROM NONMEM OUTPUT
# ============================================================================

# Disease progression parameters
THETA_SLOPE = 0.093      # Disease progression slope
THETA_POWER = 0.710      # Disease progression power

# Variance parameters (diagonal of omega via Cholesky)
VAR_DIS = 1.0           # Variance Disability (Fixed)
VAR_SLOPE = 0.2         # Variance Slope
VAR_AGE = 99.329        # Variance Age
VAR_MSD = 54.124        # Variance MSD (MS Duration)
VAR_EXNB = 0.3727       # Variance EXNB (Exacerbation number)

# Correlation parameters
COR_DIS_SLOPE = 0.11
COR_DIS_AGE = 0.2645
COR_SLOPE_AGE = 0.1175
COR_DIS_MSD = 0.2727
COR_SLOPE_MSD = 0.0896
COR_AGE_MSD = 0.4582
COR_DIS_EXNB = 0.0391
COR_SLOPE_EXNB = 0.0695
COR_AGE_EXNB = -0.0914
COR_MSD_EXNB = -0.115

# Drug effect parameters
EMAX_SYMP = 0.17        # Emax for symptomatic effect
EC50_SYMP = 408.29      # EC50 for symptomatic effect
EFF_PROT = 0.209        # Protective effect

# FREM parameters
AGE_MEAN = 38.569
MSD_MEAN = 8.6697
EXNB_MEAN = 1.354

# IRT Item parameters (boundary and slope for each functional system)
# Format: {item_name: {'boundaries': [...], 'slope': value, 'max_score': int}}
IRT_PARAMS = {
    'Pyramidal': {
        'boundaries': [-1.55, 1.248, 0.818, 1.428, 1.313],
        'slope': 3.172,
        'max_score': 5
    },
    'Cerebellar': {
        'boundaries': [-0.913, 0.963, 1.023, 1.648, 1.074],
        'slope': 2.873,
        'max_score': 5
    },
    'Brainstem': {
        'boundaries': [-0.112, 1.711, 2.000, 2.780],
        'slope': 1.038,
        'max_score': 4
    },
    'Sensory': {
        'boundaries': [-0.783, 1.186, 1.946, 2.345, 2.937, 2.353],
        'slope': 0.993,
        'max_score': 6
    },
    'Bowel/Bladder': {
        'boundaries': [-0.147, 1.887, 1.531, 2.996, 0.968],
        'slope': 1.256,
        'max_score': 5
    },
    'Visual': {
        'boundaries': [-0.037, 3.751, 2.660, 1.688, 1.705, 1.637],
        'slope': 0.440,
        'max_score': 6
    },
    'Mental': {
        'boundaries': [0.402, 1.111, 4.161, 2.890],
        'slope': 0.912,
        'max_score': 4
    },
    'Ambulation': {
        'boundaries': [1.133, 0.244, 0.225, 0.423, 0.448, 0.484, 0.139, 0.261, 0.351],
        'slope': 3.64,
        'max_score': 9
    }
}

# Objective function value
OFV = 418.626


def create_model_overview():
    """Create model structure overview infographic"""
    fig = plt.figure(figsize=(16, 12))
    fig.suptitle('IRT Model for EDSS in Multiple Sclerosis\nCladribine Treatment - Novakovic et al., 2016',
                 fontsize=16, fontweight='bold', y=0.98)

    ax = fig.add_subplot(111)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Colors
    c_disease = '#3498db'
    c_drug = '#e74c3c'
    c_irt = '#2ecc71'
    c_covariate = '#9b59b6'
    c_eta = '#f39c12'

    # Title box
    title_box = FancyBboxPatch((5, 88), 90, 10, boxstyle="round,pad=0.02",
                                facecolor='#2c3e50', edgecolor='none')
    ax.add_patch(title_box)
    ax.text(50, 93, 'Longitudinal IRT Model Structure', ha='center', va='center',
            fontsize=14, fontweight='bold', color='white')

    # Disease Latent Variable box
    disease_box = FancyBboxPatch((35, 65), 30, 15, boxstyle="round,pad=0.02",
                                  facecolor=c_disease, edgecolor='#2980b9', linewidth=2)
    ax.add_patch(disease_box)
    ax.text(50, 75, 'Disease Latent Variable', ha='center', va='center',
            fontsize=11, fontweight='bold', color='white')
    ax.text(50, 70, 'PD = P1 + (θ₁ + P2)·(t/365)^θ₂\n× (1-Ef_prot) - Ef_symp', ha='center', va='center',
            fontsize=9, color='white', family='monospace')

    # Drug Effect boxes
    # Symptomatic
    symp_box = FancyBboxPatch((5, 45), 25, 15, boxstyle="round,pad=0.02",
                               facecolor=c_drug, edgecolor='#c0392b', linewidth=2)
    ax.add_patch(symp_box)
    ax.text(17.5, 56, 'Symptomatic Effect', ha='center', va='center',
            fontsize=10, fontweight='bold', color='white')
    ax.text(17.5, 49, f'Emax = {EMAX_SYMP}\nEC50 = {EC50_SYMP}', ha='center', va='center',
            fontsize=9, color='white')

    # Protective
    prot_box = FancyBboxPatch((70, 45), 25, 15, boxstyle="round,pad=0.02",
                               facecolor=c_drug, edgecolor='#c0392b', linewidth=2)
    ax.add_patch(prot_box)
    ax.text(82.5, 56, 'Protective Effect', ha='center', va='center',
            fontsize=10, fontweight='bold', color='white')
    ax.text(82.5, 49, f'Effect = {EFF_PROT}', ha='center', va='center',
            fontsize=9, color='white')

    # Arrows from drug effects to disease
    ax.annotate('', xy=(35, 70), xytext=(30, 55),
                arrowprops=dict(arrowstyle='->', color=c_drug, lw=2))
    ax.annotate('', xy=(65, 70), xytext=(70, 55),
                arrowprops=dict(arrowstyle='->', color=c_drug, lw=2))

    # Random Effects box (left side)
    eta_box = FancyBboxPatch((5, 65), 25, 20, boxstyle="round,pad=0.02",
                              facecolor=c_eta, edgecolor='#d68910', linewidth=2)
    ax.add_patch(eta_box)
    ax.text(17.5, 82, 'Random Effects', ha='center', va='center',
            fontsize=10, fontweight='bold', color='white')
    ax.text(17.5, 73, 'η₁: Disability\nη₂: Slope\nη₃: Age\nη₄: MSD\nη₅: EXNB\nη₆: Drug Effect',
            ha='center', va='center', fontsize=8, color='white')

    # Arrow from random effects to disease
    ax.annotate('', xy=(35, 72), xytext=(30, 75),
                arrowprops=dict(arrowstyle='->', color=c_eta, lw=2))

    # Covariates box (right side)
    cov_box = FancyBboxPatch((70, 65), 25, 20, boxstyle="round,pad=0.02",
                              facecolor=c_covariate, edgecolor='#7d3c98', linewidth=2)
    ax.add_patch(cov_box)
    ax.text(82.5, 82, 'FREM Covariates', ha='center', va='center',
            fontsize=10, fontweight='bold', color='white')
    ax.text(82.5, 73, f'Age: {AGE_MEAN:.1f} yrs\nMSD: {MSD_MEAN:.1f} yrs\nEXNB: {EXNB_MEAN:.2f}',
            ha='center', va='center', fontsize=9, color='white')

    # Arrow from covariates to disease
    ax.annotate('', xy=(65, 72), xytext=(70, 75),
                arrowprops=dict(arrowstyle='->', color=c_covariate, lw=2))

    # IRT Functional Systems (bottom)
    fs_names = ['Pyramidal', 'Cerebellar', 'Brainstem', 'Sensory',
                'Bowel/\nBladder', 'Visual', 'Mental', 'Ambulation']
    fs_max = [5, 5, 4, 6, 5, 6, 4, 9]

    start_x = 5
    box_width = 10.5
    gap = 1

    for i, (name, max_s) in enumerate(zip(fs_names, fs_max)):
        x = start_x + i * (box_width + gap)
        fs_box = FancyBboxPatch((x, 15), box_width, 22, boxstyle="round,pad=0.02",
                                 facecolor=c_irt, edgecolor='#27ae60', linewidth=2)
        ax.add_patch(fs_box)
        ax.text(x + box_width/2, 32, name, ha='center', va='center',
                fontsize=8, fontweight='bold', color='white')
        ax.text(x + box_width/2, 21, f'Score: 0-{max_s}', ha='center', va='center',
                fontsize=8, color='white')

        # Arrow from disease to FS
        ax.annotate('', xy=(x + box_width/2, 37), xytext=(50, 65),
                    arrowprops=dict(arrowstyle='->', color='#34495e', lw=1, alpha=0.5))

    # IRT box label
    irt_label = FancyBboxPatch((25, 38), 50, 8, boxstyle="round,pad=0.02",
                                facecolor='white', edgecolor=c_irt, linewidth=2)
    ax.add_patch(irt_label)
    ax.text(50, 42, 'Item Response Theory (Graded Response Model)', ha='center', va='center',
            fontsize=10, fontweight='bold', color=c_irt)

    # Legend box
    legend_box = FancyBboxPatch((5, 2), 90, 10, boxstyle="round,pad=0.02",
                                 facecolor='#ecf0f1', edgecolor='#bdc3c7', linewidth=1)
    ax.add_patch(legend_box)

    # Legend items
    legend_items = [
        (c_disease, 'Latent Disease'),
        (c_drug, 'Drug Effect'),
        (c_irt, 'IRT Items'),
        (c_eta, 'Random Effects'),
        (c_covariate, 'Covariates')
    ]

    for i, (color, label) in enumerate(legend_items):
        x = 10 + i * 18
        legend_circle = Circle((x, 7), 1.5, facecolor=color, edgecolor='white')
        ax.add_patch(legend_circle)
        ax.text(x + 3, 7, label, ha='left', va='center', fontsize=9)

    plt.tight_layout(rect=[0, 0, 1, 0.96])
    plt.savefig('infographic_01_model_structure.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print("Created: infographic_01_model_structure.png")


def create_parameter_table():
    """Create parameter estimates visualization"""
    fig = plt.figure(figsize=(14, 16))
    fig.suptitle('Parameter Estimates Summary\nIRT Model - Cladribine in Multiple Sclerosis',
                 fontsize=14, fontweight='bold', y=0.99)

    gs = GridSpec(4, 2, figure=fig, hspace=0.3, wspace=0.25)

    # 1. Disease Progression Parameters
    ax1 = fig.add_subplot(gs[0, 0])
    ax1.axis('off')
    ax1.set_title('Disease Progression Parameters', fontsize=12, fontweight='bold',
                  color='#2c3e50', pad=10)

    dp_params = [
        ('Slope (θ₁)', THETA_SLOPE, 'Disease progression rate'),
        ('Power (θ₂)', THETA_POWER, 'Time-course exponent'),
    ]

    y_pos = 0.85
    for name, value, desc in dp_params:
        ax1.text(0.05, y_pos, name, fontsize=11, fontweight='bold', transform=ax1.transAxes)
        ax1.text(0.55, y_pos, f'{value:.4f}', fontsize=11, transform=ax1.transAxes, family='monospace')
        ax1.text(0.05, y_pos - 0.08, desc, fontsize=9, color='gray', transform=ax1.transAxes)
        y_pos -= 0.25

    # Add equation
    ax1.text(0.5, 0.1, r'$PD = P_1 + (\theta_1 + P_2) \cdot (t/365)^{\theta_2}$',
             fontsize=11, transform=ax1.transAxes, ha='center',
             bbox=dict(boxstyle='round', facecolor='#e8f4fd', edgecolor='#3498db'))

    # 2. Drug Effect Parameters
    ax2 = fig.add_subplot(gs[0, 1])
    ax2.axis('off')
    ax2.set_title('Drug Effect Parameters', fontsize=12, fontweight='bold',
                  color='#c0392b', pad=10)

    drug_params = [
        ('Emax (symptomatic)', EMAX_SYMP, 'Maximum symptomatic effect'),
        ('EC50 (symptomatic)', EC50_SYMP, 'Half-maximal exposure'),
        ('Protective Effect', EFF_PROT, 'Disease-modifying fraction'),
    ]

    y_pos = 0.85
    for name, value, desc in drug_params:
        ax2.text(0.05, y_pos, name, fontsize=11, fontweight='bold', transform=ax2.transAxes)
        ax2.text(0.65, y_pos, f'{value:.3f}', fontsize=11, transform=ax2.transAxes, family='monospace')
        ax2.text(0.05, y_pos - 0.08, desc, fontsize=9, color='gray', transform=ax2.transAxes)
        y_pos -= 0.22

    # 3. Variance Parameters
    ax3 = fig.add_subplot(gs[1, 0])
    variance_names = ['Disability\n(Fixed)', 'Slope', 'Age', 'MSD', 'EXNB']
    variance_values = [VAR_DIS, VAR_SLOPE, VAR_AGE, VAR_MSD, VAR_EXNB]

    colors = ['#95a5a6' if v == 1.0 and n.startswith('Disability') else '#3498db'
              for v, n in zip(variance_values, variance_names)]

    bars = ax3.bar(variance_names, variance_values, color=colors, edgecolor='#2c3e50', linewidth=1.5)
    ax3.set_ylabel('Variance', fontsize=11)
    ax3.set_title('Random Effect Variances', fontsize=12, fontweight='bold', color='#2c3e50')
    ax3.set_ylim(0, max(variance_values) * 1.15)

    for bar, val in zip(bars, variance_values):
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2., height + 1,
                f'{val:.2f}', ha='center', va='bottom', fontsize=9, fontweight='bold')

    ax3.spines['top'].set_visible(False)
    ax3.spines['right'].set_visible(False)

    # 4. FREM Covariate Means
    ax4 = fig.add_subplot(gs[1, 1])
    cov_names = ['Age (years)', 'MSD (years)', 'EXNB']
    cov_values = [AGE_MEAN, MSD_MEAN, EXNB_MEAN]

    bars = ax4.barh(cov_names, cov_values, color='#9b59b6', edgecolor='#7d3c98', linewidth=1.5)
    ax4.set_xlabel('Mean Value', fontsize=11)
    ax4.set_title('FREM Covariate Population Means', fontsize=12, fontweight='bold', color='#7d3c98')

    for bar, val in zip(bars, cov_values):
        width = bar.get_width()
        ax4.text(width + 0.5, bar.get_y() + bar.get_height()/2.,
                f'{val:.2f}', ha='left', va='center', fontsize=10, fontweight='bold')

    ax4.spines['top'].set_visible(False)
    ax4.spines['right'].set_visible(False)

    # 5. IRT Item Slopes
    ax5 = fig.add_subplot(gs[2, :])
    item_names = list(IRT_PARAMS.keys())
    slopes = [IRT_PARAMS[item]['slope'] for item in item_names]
    max_scores = [IRT_PARAMS[item]['max_score'] for item in item_names]

    x = np.arange(len(item_names))
    width = 0.35

    bars1 = ax5.bar(x - width/2, slopes, width, label='Discrimination (slope)',
                    color='#2ecc71', edgecolor='#27ae60', linewidth=1.5)

    ax5_twin = ax5.twinx()
    bars2 = ax5_twin.bar(x + width/2, max_scores, width, label='Max Score',
                         color='#e74c3c', edgecolor='#c0392b', linewidth=1.5, alpha=0.7)

    ax5.set_xlabel('Functional System', fontsize=11)
    ax5.set_ylabel('Discrimination Parameter', fontsize=11, color='#27ae60')
    ax5_twin.set_ylabel('Maximum Score', fontsize=11, color='#c0392b')
    ax5.set_title('IRT Item Parameters by Functional System', fontsize=12, fontweight='bold')
    ax5.set_xticks(x)
    ax5.set_xticklabels(item_names, rotation=45, ha='right')

    for bar, val in zip(bars1, slopes):
        height = bar.get_height()
        ax5.text(bar.get_x() + bar.get_width()/2., height + 0.05,
                f'{val:.2f}', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#27ae60')

    ax5.legend(loc='upper left')
    ax5_twin.legend(loc='upper right')
    ax5.spines['top'].set_visible(False)

    # 6. Model Fit Statistics
    ax6 = fig.add_subplot(gs[3, :])
    ax6.axis('off')
    ax6.set_title('Model Fit & Estimation Summary', fontsize=12, fontweight='bold', pad=10)

    # Create info boxes
    info_boxes = [
        ('Objective Function', f'{OFV:.3f}', '#3498db'),
        ('Estimation Method', 'LAPLACIAN', '#2ecc71'),
        ('N Observations', '241', '#e74c3c'),
        ('N Subjects', '3', '#9b59b6'),
        ('N Thetas', '75', '#f39c12'),
        ('N Omegas', '6', '#1abc9c'),
    ]

    for i, (label, value, color) in enumerate(info_boxes):
        x = 0.08 + (i % 3) * 0.32
        y = 0.6 if i < 3 else 0.2

        box = FancyBboxPatch((x - 0.02, y - 0.15), 0.28, 0.3, transform=ax6.transAxes,
                             boxstyle="round,pad=0.02", facecolor=color, alpha=0.15,
                             edgecolor=color, linewidth=2)
        ax6.add_patch(box)
        ax6.text(x + 0.12, y + 0.08, label, fontsize=10, transform=ax6.transAxes,
                ha='center', va='center', fontweight='bold', color=color)
        ax6.text(x + 0.12, y - 0.05, value, fontsize=14, transform=ax6.transAxes,
                ha='center', va='center', fontweight='bold', color='#2c3e50')

    plt.tight_layout(rect=[0, 0, 1, 0.97])
    plt.savefig('infographic_02_parameters.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print("Created: infographic_02_parameters.png")


def create_correlation_heatmap():
    """Create correlation matrix heatmap for random effects"""
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.suptitle('Random Effects Correlation Matrix (Cholesky Parameterization)',
                 fontsize=14, fontweight='bold')

    # Build correlation matrix
    labels = ['Disability', 'Slope', 'Age', 'MSD', 'EXNB']
    n = len(labels)
    corr_matrix = np.eye(n)

    # Fill in correlations
    corr_matrix[0, 1] = corr_matrix[1, 0] = COR_DIS_SLOPE
    corr_matrix[0, 2] = corr_matrix[2, 0] = COR_DIS_AGE
    corr_matrix[0, 3] = corr_matrix[3, 0] = COR_DIS_MSD
    corr_matrix[0, 4] = corr_matrix[4, 0] = COR_DIS_EXNB
    corr_matrix[1, 2] = corr_matrix[2, 1] = COR_SLOPE_AGE
    corr_matrix[1, 3] = corr_matrix[3, 1] = COR_SLOPE_MSD
    corr_matrix[1, 4] = corr_matrix[4, 1] = COR_SLOPE_EXNB
    corr_matrix[2, 3] = corr_matrix[3, 2] = COR_AGE_MSD
    corr_matrix[2, 4] = corr_matrix[4, 2] = COR_AGE_EXNB
    corr_matrix[3, 4] = corr_matrix[4, 3] = COR_MSD_EXNB

    # Create custom colormap
    colors = ['#c0392b', '#e74c3c', '#f5b7b1', 'white', '#aed6f1', '#3498db', '#1a5276']
    cmap = LinearSegmentedColormap.from_list('corr', colors, N=256)

    # Plot heatmap
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool), k=1)

    sns.heatmap(corr_matrix, annot=True, fmt='.3f', cmap=cmap, center=0,
                xticklabels=labels, yticklabels=labels,
                vmin=-1, vmax=1, linewidths=0.5, linecolor='white',
                cbar_kws={'label': 'Correlation', 'shrink': 0.8},
                annot_kws={'size': 12, 'weight': 'bold'},
                ax=ax)

    ax.set_title('', fontsize=12)

    # Rotate labels
    plt.xticks(rotation=45, ha='right')
    plt.yticks(rotation=0)

    # Add interpretation box
    textstr = '\n'.join([
        'Key Correlations:',
        f'• Age-MSD: {COR_AGE_MSD:.3f} (moderate positive)',
        f'• Disability-MSD: {COR_DIS_MSD:.3f} (weak positive)',
        f'• Disability-Age: {COR_DIS_AGE:.3f} (weak positive)',
        f'• MSD-EXNB: {COR_MSD_EXNB:.3f} (weak negative)'
    ])
    props = dict(boxstyle='round', facecolor='wheat', alpha=0.8)
    ax.text(1.02, 0.5, textstr, transform=ax.transAxes, fontsize=10,
            verticalalignment='center', bbox=props)

    plt.tight_layout()
    plt.savefig('infographic_03_correlation_matrix.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print("Created: infographic_03_correlation_matrix.png")


def create_icc_curves():
    """Create Item Characteristic Curves for IRT items"""
    fig, axes = plt.subplots(2, 4, figsize=(16, 10))
    fig.suptitle('Item Characteristic Curves (ICC)\nProbability of Scoring ≥ k at Each Disability Level',
                 fontsize=14, fontweight='bold')

    theta_range = np.linspace(-3, 6, 200)
    colors = plt.cm.viridis(np.linspace(0, 0.9, 10))

    for idx, (item_name, params) in enumerate(IRT_PARAMS.items()):
        ax = axes.flatten()[idx]

        boundaries = params['boundaries']
        slope = params['slope']
        max_score = params['max_score']

        # Calculate cumulative boundary positions
        cum_boundaries = [boundaries[0]]
        for i in range(1, len(boundaries)):
            cum_boundaries.append(cum_boundaries[-1] + boundaries[i])

        # Plot ICC for each score level
        for k, b in enumerate(cum_boundaries):
            logit = slope * (theta_range - b)
            prob = 1 / (1 + np.exp(-logit))
            ax.plot(theta_range, prob, color=colors[k], linewidth=2,
                   label=f'P(X≥{k+1})')

        ax.set_xlabel('Disability (θ)', fontsize=10)
        ax.set_ylabel('Probability', fontsize=10)
        ax.set_title(f'{item_name}\n(slope={slope:.2f})', fontsize=11, fontweight='bold')
        ax.set_ylim(0, 1)
        ax.set_xlim(-3, 6)
        ax.grid(True, alpha=0.3)
        ax.legend(loc='lower right', fontsize=7, ncol=2)
        ax.axhline(y=0.5, color='gray', linestyle='--', alpha=0.5)

        # Add vertical line at mean
        ax.axvline(x=0, color='red', linestyle=':', alpha=0.5, label='Mean disability')

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig('infographic_04_icc_curves.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print("Created: infographic_04_icc_curves.png")


def create_drug_effect_plot():
    """Create drug effect visualization"""
    fig = plt.figure(figsize=(14, 10))
    gs = GridSpec(2, 2, figure=fig, hspace=0.3, wspace=0.3)
    fig.suptitle('Cladribine Drug Effects on Disease Progression',
                 fontsize=14, fontweight='bold')

    # 1. Symptomatic Effect (Emax model)
    ax1 = fig.add_subplot(gs[0, 0])
    exposure = np.linspace(0, 2000, 200)
    symp_effect = EMAX_SYMP * exposure / (exposure + EC50_SYMP)

    ax1.plot(exposure, symp_effect, 'b-', linewidth=3, label='Symptomatic Effect')
    ax1.axhline(y=EMAX_SYMP, color='gray', linestyle='--', alpha=0.7, label=f'Emax = {EMAX_SYMP}')
    ax1.axvline(x=EC50_SYMP, color='red', linestyle=':', alpha=0.7, label=f'EC50 = {EC50_SYMP}')
    ax1.axhline(y=EMAX_SYMP/2, color='green', linestyle=':', alpha=0.5)

    ax1.fill_between(exposure, symp_effect, alpha=0.2, color='blue')
    ax1.set_xlabel('Surrogate Exposure (CD × 104.5 / CrCL)', fontsize=11)
    ax1.set_ylabel('Symptomatic Effect', fontsize=11)
    ax1.set_title('Exposure-Response: Symptomatic Effect', fontsize=12, fontweight='bold')
    ax1.legend(loc='lower right')
    ax1.set_xlim(0, 2000)
    ax1.set_ylim(0, EMAX_SYMP * 1.1)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

    # Add annotation
    ax1.annotate(f'EC50 = {EC50_SYMP}', xy=(EC50_SYMP, EMAX_SYMP/2),
                xytext=(EC50_SYMP + 300, EMAX_SYMP/2 + 0.02),
                arrowprops=dict(arrowstyle='->', color='red'),
                fontsize=10, color='red')

    # 2. Disease Progression with/without Treatment
    ax2 = fig.add_subplot(gs[0, 1])
    time_years = np.linspace(0, 5, 100)
    time_days = time_years * 365

    # Placebo progression
    prog_placebo = THETA_SLOPE * (time_days / 365) ** THETA_POWER

    # With protective effect only
    prog_protective = THETA_SLOPE * (time_days / 365) ** THETA_POWER * (1 - EFF_PROT)

    # With both effects (assuming median exposure)
    median_exposure = EC50_SYMP
    symp_at_median = EMAX_SYMP * median_exposure / (median_exposure + EC50_SYMP)
    prog_full = prog_protective - symp_at_median

    ax2.plot(time_years, prog_placebo, 'r-', linewidth=2.5, label='Placebo')
    ax2.plot(time_years, prog_protective, 'orange', linewidth=2.5,
             label=f'Protective Effect Only ({EFF_PROT*100:.0f}%)')
    ax2.plot(time_years, prog_full, 'g-', linewidth=2.5,
             label='Protective + Symptomatic')

    ax2.fill_between(time_years, prog_placebo, prog_full, alpha=0.2, color='green',
                     label='Treatment Benefit')

    ax2.set_xlabel('Time (years)', fontsize=11)
    ax2.set_ylabel('Disease Progression (PD)', fontsize=11)
    ax2.set_title('Disease Progression Over Time', fontsize=12, fontweight='bold')
    ax2.legend(loc='upper left')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.set_xlim(0, 5)

    # 3. Protective Effect Impact
    ax3 = fig.add_subplot(gs[1, 0])

    categories = ['Slope\n(No Treatment)', 'Slope\n(With Protection)']
    values = [THETA_SLOPE, THETA_SLOPE * (1 - EFF_PROT)]
    colors = ['#e74c3c', '#2ecc71']

    bars = ax3.bar(categories, values, color=colors, edgecolor='#2c3e50', linewidth=2)

    for bar, val in zip(bars, values):
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2., height + 0.002,
                f'{val:.4f}', ha='center', va='bottom', fontsize=11, fontweight='bold')

    ax3.set_ylabel('Progression Rate', fontsize=11)
    ax3.set_title(f'Protective Effect: {EFF_PROT*100:.1f}% Reduction in Progression Rate',
                  fontsize=12, fontweight='bold')
    ax3.spines['top'].set_visible(False)
    ax3.spines['right'].set_visible(False)

    # Add reduction arrow
    ax3.annotate('', xy=(1, values[1]), xytext=(0, values[0]),
                arrowprops=dict(arrowstyle='->', color='blue', lw=2))
    ax3.text(0.5, (values[0] + values[1])/2 + 0.01, f'-{EFF_PROT*100:.1f}%',
             ha='center', fontsize=12, fontweight='bold', color='blue')

    # 4. Combined Effect Summary
    ax4 = fig.add_subplot(gs[1, 1])
    ax4.axis('off')

    # Create summary box
    summary_text = """
    Drug Effect Components:

    1. SYMPTOMATIC EFFECT
       • Direct reduction in disability latent variable
       • Exposure-dependent (Emax model)
       • Emax = {:.2f}, EC50 = {:.1f}
       • Effect formula: Emax × Exp / (Exp + EC50)

    2. PROTECTIVE (DISEASE-MODIFYING) EFFECT
       • Reduces rate of disease progression
       • Exposure-independent (binary)
       • Reduces slope by {:.1f}%
       • Effect formula: (1 - {:.3f}) × slope

    Model Equation:
    PD = P1 + (θ₁ + P2)·(t/365)^θ₂ × (1-Ef_prot) - Ef_symp

    Where Ef_symp and Ef_prot are active when TRT ≥ 1 and t > 0
    """.format(EMAX_SYMP, EC50_SYMP, EFF_PROT*100, EFF_PROT)

    props = dict(boxstyle='round', facecolor='#f8f9fa', edgecolor='#2c3e50', linewidth=2)
    ax4.text(0.5, 0.5, summary_text, transform=ax4.transAxes, fontsize=10,
             verticalalignment='center', horizontalalignment='center', bbox=props,
             family='monospace')
    ax4.set_title('Drug Effect Summary', fontsize=12, fontweight='bold')

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig('infographic_05_drug_effects.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print("Created: infographic_05_drug_effects.png")


def create_irt_summary():
    """Create IRT item parameter summary visualization"""
    fig = plt.figure(figsize=(16, 12))
    fig.suptitle('IRT Graded Response Model - Item Parameter Summary\nEDSS Functional System Subscores',
                 fontsize=14, fontweight='bold')

    gs = GridSpec(3, 3, figure=fig, hspace=0.35, wspace=0.3)

    # 1. Item Discrimination Comparison (top left, spans 2 cols)
    ax1 = fig.add_subplot(gs[0, :2])

    items = list(IRT_PARAMS.keys())
    slopes = [IRT_PARAMS[item]['slope'] for item in items]

    colors = plt.cm.RdYlGn([s/max(slopes) for s in slopes])
    bars = ax1.barh(items, slopes, color=colors, edgecolor='#2c3e50', linewidth=1.5)

    ax1.set_xlabel('Discrimination Parameter (Slope)', fontsize=11)
    ax1.set_title('Item Discrimination: Higher = Better Differentiation', fontsize=12, fontweight='bold')
    ax1.axvline(x=np.mean(slopes), color='blue', linestyle='--', label=f'Mean: {np.mean(slopes):.2f}')

    for bar, val in zip(bars, slopes):
        width = bar.get_width()
        ax1.text(width + 0.05, bar.get_y() + bar.get_height()/2.,
                f'{val:.2f}', ha='left', va='center', fontsize=10, fontweight='bold')

    ax1.legend(loc='lower right')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)

    # 2. Interpretation Guide (top right)
    ax2 = fig.add_subplot(gs[0, 2])
    ax2.axis('off')

    guide_text = """
    IRT Discrimination Interpretation:

    Slope < 1.0:  Low discrimination
    1.0 - 1.5:    Moderate discrimination
    1.5 - 2.5:    High discrimination
    > 2.5:        Very high discrimination

    High discrimination items:
    • Ambulation (3.64)
    • Pyramidal (3.17)
    • Cerebellar (2.87)

    These items are most informative
    for distinguishing disability levels.
    """

    props = dict(boxstyle='round', facecolor='#e8f6e8', edgecolor='#27ae60', linewidth=2)
    ax2.text(0.5, 0.5, guide_text, transform=ax2.transAxes, fontsize=10,
             verticalalignment='center', horizontalalignment='center', bbox=props)

    # 3-4. Item Information Curves
    ax3 = fig.add_subplot(gs[1, :])
    theta_range = np.linspace(-3, 6, 200)

    colors = plt.cm.tab10(np.linspace(0, 1, 8))

    for idx, (item_name, params) in enumerate(IRT_PARAMS.items()):
        slope = params['slope']
        boundaries = params['boundaries']

        # Calculate total item information
        cum_boundaries = [boundaries[0]]
        for i in range(1, len(boundaries)):
            cum_boundaries.append(cum_boundaries[-1] + boundaries[i])

        total_info = np.zeros_like(theta_range)
        for b in cum_boundaries:
            logit = slope * (theta_range - b)
            p = 1 / (1 + np.exp(-logit))
            info = (slope ** 2) * p * (1 - p)
            total_info += info

        ax3.plot(theta_range, total_info, color=colors[idx], linewidth=2, label=item_name)

    ax3.set_xlabel('Disability Level (θ)', fontsize=11)
    ax3.set_ylabel('Information', fontsize=11)
    ax3.set_title('Item Information Functions\nPeak shows where item is most informative', fontsize=12, fontweight='bold')
    ax3.legend(loc='upper right', ncol=4)
    ax3.set_xlim(-3, 6)
    ax3.axvline(x=0, color='gray', linestyle=':', alpha=0.5)
    ax3.spines['top'].set_visible(False)
    ax3.spines['right'].set_visible(False)
    ax3.grid(True, alpha=0.3)

    # 5. Total Test Information
    ax4 = fig.add_subplot(gs[2, 0])

    total_test_info = np.zeros_like(theta_range)
    for item_name, params in IRT_PARAMS.items():
        slope = params['slope']
        boundaries = params['boundaries']

        cum_boundaries = [boundaries[0]]
        for i in range(1, len(boundaries)):
            cum_boundaries.append(cum_boundaries[-1] + boundaries[i])

        for b in cum_boundaries:
            logit = slope * (theta_range - b)
            p = 1 / (1 + np.exp(-logit))
            info = (slope ** 2) * p * (1 - p)
            total_test_info += info

    ax4.fill_between(theta_range, total_test_info, alpha=0.3, color='blue')
    ax4.plot(theta_range, total_test_info, 'b-', linewidth=2)
    ax4.set_xlabel('Disability Level (θ)', fontsize=11)
    ax4.set_ylabel('Total Information', fontsize=11)
    ax4.set_title('Test Information Function', fontsize=12, fontweight='bold')
    ax4.axvline(x=0, color='gray', linestyle=':', alpha=0.5)
    ax4.spines['top'].set_visible(False)
    ax4.spines['right'].set_visible(False)

    # Standard error of measurement
    se = 1 / np.sqrt(total_test_info + 1e-10)
    ax4_twin = ax4.twinx()
    ax4_twin.plot(theta_range, se, 'r--', linewidth=1.5, alpha=0.7, label='SE')
    ax4_twin.set_ylabel('Standard Error', fontsize=10, color='red')
    ax4_twin.tick_params(axis='y', labelcolor='red')
    ax4_twin.set_ylim(0, 2)

    # 6. Boundary Locations
    ax5 = fig.add_subplot(gs[2, 1:])

    y_pos = np.arange(len(items))
    colors = plt.cm.Set3(np.linspace(0, 1, 10))

    for idx, (item_name, params) in enumerate(IRT_PARAMS.items()):
        boundaries = params['boundaries']
        cum_boundaries = [boundaries[0]]
        for i in range(1, len(boundaries)):
            cum_boundaries.append(cum_boundaries[-1] + boundaries[i])

        for k, b in enumerate(cum_boundaries):
            ax5.scatter(b, idx, c=[colors[k]], s=100, edgecolors='black', linewidths=1, zorder=3)
            if k == 0:
                ax5.plot([b - 0.1, b], [idx, idx], 'k-', linewidth=1)

        # Connect boundaries with lines
        ax5.plot(cum_boundaries, [idx] * len(cum_boundaries), 'k-', linewidth=1, alpha=0.5)

    ax5.set_yticks(y_pos)
    ax5.set_yticklabels(items)
    ax5.set_xlabel('Disability Level (θ)', fontsize=11)
    ax5.set_title('Boundary Locations (Thresholds)\nPoints show where P(X≥k) = 0.5', fontsize=12, fontweight='bold')
    ax5.axvline(x=0, color='red', linestyle='--', alpha=0.5, label='Mean disability')
    ax5.grid(True, axis='x', alpha=0.3)
    ax5.legend(loc='lower right')
    ax5.spines['top'].set_visible(False)
    ax5.spines['right'].set_visible(False)

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig('infographic_06_irt_summary.png', dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print("Created: infographic_06_irt_summary.png")


def main():
    """Generate all infographics"""
    print("="*60)
    print("NONMEM IRT Model Infographics Generator")
    print("Model: Cladribine in Multiple Sclerosis - EDSS IRT")
    print("Reference: Novakovic et al., 2016")
    print("="*60)
    print()

    print("Generating infographics...")
    print()

    create_model_overview()
    create_parameter_table()
    create_correlation_heatmap()
    create_icc_curves()
    create_drug_effect_plot()
    create_irt_summary()

    print()
    print("="*60)
    print("All infographics generated successfully!")
    print("="*60)
    print()
    print("Generated files:")
    print("  1. infographic_01_model_structure.png")
    print("  2. infographic_02_parameters.png")
    print("  3. infographic_03_correlation_matrix.png")
    print("  4. infographic_04_icc_curves.png")
    print("  5. infographic_05_drug_effects.png")
    print("  6. infographic_06_irt_summary.png")


if __name__ == '__main__':
    main()
