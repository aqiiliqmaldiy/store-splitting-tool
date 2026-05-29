import pandas as pd
import numpy as np

# =========================================================
# CONFIGURATION
# =========================================================

input_file = "Store_Splitting_Summary.xlsx"     # <-- Change to your file name
output_file = "output.xlsx"

# Data columns (D to O)
data_columns = [
    "F1", "F2", "F3", "F4", "F5", "F6",
    "G1", "G2", "G3", "G4", "G5", "G6"
]

# =========================================================
# METRIC FUNCTIONS
# =========================================================

def calculate_balance_score(y):
    """
    Balance Score:
    1 = perfectly balanced (left sum == right sum)
    0 = extremely imbalanced (all mass on one side)
    """
    total = sum(y)

    if total == 0:
        return 0.0

    n = len(y)
    mid = n // 2

    left_sum = sum(y[:mid])
    right_sum = total - left_sum

    bi = (left_sum - right_sum) / total
    score = 1.0 - abs(bi)

    return round(max(0.0, min(1.0, score)), 6)


def calculate_symmetry_score(y):
    """
    Symmetry Score:
    1 = perfectly symmetric
    0 = highly asymmetric
    """

    total = sum(y)

    if total == 0:
        return 0

    n = len(y)
    mid = n // 2

    left = y[:mid]

    if n % 2 == 0:
        right = y[mid:]
    else:
        right = y[mid + 1:]

    right_rev = right[::-1]

    m = min(len(left), len(right_rev))

    abs_diff = 0

    for i in range(m):
        abs_diff += abs((left[i] / total) - (right_rev[i] / total))

    symmetry = max(0.0, min(1.0, 1.0 - abs_diff))

    return round(symmetry, 6)


def calculate_center_score(y):
    """
    Center Score:
    1 = concentrated at center
    0 = concentrated at edges
    """

    total = sum(y)

    if total == 0:
        return 0

    n = len(y)

    center_idx = (n - 1) / 2.0
    max_dist = max(center_idx, (n - 1) - center_idx)

    weighted_sum = 0

    for idx, value in enumerate(y):

        weight = 1.0 - abs(idx - center_idx) / max_dist

        weighted_sum += value * weight

    center_score = weighted_sum / total

    return round(center_score, 6)


def calculate_final_score(balance_score, symmetry_score, center_score):
    """
    Final Score:
    Average of Balance Score, Symmetry Score, Center Score
    """

    final_score = (
        balance_score +
        symmetry_score +
        center_score
    ) / 3

    return round(final_score, 6)

# =========================================================
# READ EXCEL
# =========================================================

df = pd.read_excel(input_file)

# =========================================================
# CALCULATE SCORES
# =========================================================

balance_scores = []
symmetry_scores = []
center_scores = []
final_scores = []

for _, row in df.iterrows():

    y = row[data_columns].fillna(0).astype(float).tolist()

    # -------------------------
    # Calculate metrics
    # -------------------------

    balance_score = calculate_balance_score(y)

    symmetry_score = calculate_symmetry_score(y)

    center_score = calculate_center_score(y)

    final_score = calculate_final_score(
        balance_score,
        symmetry_score,
        center_score
    )

    # -------------------------
    # Save results
    # -------------------------

    balance_scores.append(balance_score)
    symmetry_scores.append(symmetry_score)
    center_scores.append(center_score)
    final_scores.append(final_score)

# =========================================================
# WRITE RESULTS
# =========================================================

df["Balance Score"] = balance_scores
df["Symmetry Score"] = symmetry_scores
df["Center Score"] = center_scores
df["Final Score"] = final_scores

# =========================================================
# EXPORT NEW EXCEL
# =========================================================

df.to_excel(output_file, index=False)

print("Completed!")
print(f"Output saved to: {output_file}")