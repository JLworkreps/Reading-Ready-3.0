"""
Machine Learning Profiling Module
----------------------------------
Implements the two analytical components described in the project scope:

1. Literacy pattern clustering (K-Means vs DBSCAN) over word recognition
   accuracy, comprehension average, and reading speed -- the higher-performing
   model (by silhouette score) is deployed.

2. Academic risk prediction (Linear Regression) -- predicts a student's next
   composite reading score from their assessment history and flags students
   likely to fall into/remain in the Frustration level.
"""
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans, DBSCAN
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score
from sklearn.linear_model import LinearRegression

FRUSTRATION_COMPOSITE_THRESHOLD = 65.0  # below this composite score => at risk


def _composite_score(row):
    """Simple weighted composite: 40% word recognition, 60% comprehension avg."""
    return 0.4 * row["word_recognition_accuracy"] + 0.6 * row["comprehension_avg"]


def build_feature_frame(assessments):
    """assessments: list of Assessment ORM objects (latest term per student ideally)."""
    rows = []
    for a in assessments:
        rows.append({
            "student_id": a.student_id,
            "word_recognition_accuracy": a.word_recognition_accuracy,
            "comprehension_avg": a.comprehension_avg,
            "reading_speed_wpm": a.reading_speed_wpm,
            "reading_level": a.reading_level,
            "term": a.term,
            "school_year": a.school_year,
        })
    return pd.DataFrame(rows)


def cluster_literacy_patterns(df):
    """
    Runs K-Means and DBSCAN on [word_recognition_accuracy, comprehension_avg,
    reading_speed_wpm], compares silhouette scores, and returns cluster labels
    from whichever model performs better, plus metadata about the choice.
    """
    if df.empty or len(df) < 4:
        return df.assign(cluster=-1), {"model_used": "none", "reason": "insufficient data (need >= 4 students)"}

    features = df[["word_recognition_accuracy", "comprehension_avg", "reading_speed_wpm"]].values
    X = StandardScaler().fit_transform(features)

    results = {}

    # --- K-Means (try k=2..4, pick best silhouette) ---
    best_km = None
    best_km_score = -1
    for k in range(2, min(5, len(df))):
        try:
            km = KMeans(n_clusters=k, random_state=42, n_init=10).fit(X)
            if len(set(km.labels_)) < 2:
                continue
            score = silhouette_score(X, km.labels_)
            if score > best_km_score:
                best_km_score = score
                best_km = km
        except Exception:
            continue
    results["kmeans"] = (best_km, best_km_score)

    # --- DBSCAN ---
    try:
        dbscan = DBSCAN(eps=0.9, min_samples=2).fit(X)
        labels = dbscan.labels_
        if len(set(labels)) > 1 and len(set(labels)) < len(df):
            db_score = silhouette_score(X, labels)
        else:
            db_score = -1
    except Exception:
        dbscan, db_score = None, -1
    results["dbscan"] = (dbscan, db_score)

    # --- pick the better model ---
    km_model, km_score = results["kmeans"]
    db_model, db_score = results["dbscan"]

    if km_score >= db_score and km_model is not None:
        chosen_model, chosen_score, model_name = km_model, km_score, "K-Means"
        labels = chosen_model.labels_
    elif db_model is not None:
        chosen_model, chosen_score, model_name = db_model, db_score, "DBSCAN"
        labels = chosen_model.labels_
    else:
        return df.assign(cluster=-1), {"model_used": "none", "reason": "no valid clustering found"}

    df = df.copy()
    df["cluster"] = labels

    meta = {
        "model_used": model_name,
        "silhouette_score": round(float(chosen_score), 3),
        "kmeans_silhouette": round(float(km_score), 3) if km_model is not None else None,
        "dbscan_silhouette": round(float(db_score), 3) if db_model is not None else None,
        "n_clusters": len(set(labels)) - (1 if -1 in labels else 0),
    }
    return df, meta


def predict_academic_risk(student_history_df):
    """
    student_history_df: DataFrame with columns [student_id, term_order, composite]
    for ONE student, ordered chronologically (pre -> mid -> post across years).
    Fits a simple linear regression over term index -> composite score and
    extrapolates the next term. Requires >= 2 historical points; otherwise
    falls back to the latest known composite.
    """
    if student_history_df.empty:
        return None

    df = student_history_df.sort_values("term_order")
    if len(df) < 2:
        latest = df.iloc[-1]["composite"]
        return {
            "predicted_next_composite": round(float(latest), 2),
            "trend": "insufficient_history",
            "at_risk": latest < FRUSTRATION_COMPOSITE_THRESHOLD,
        }

    X = df["term_order"].values.reshape(-1, 1)
    y = df["composite"].values
    model = LinearRegression().fit(X, y)
    next_term = df["term_order"].max() + 1
    predicted = model.predict([[next_term]])[0]
    predicted = float(np.clip(predicted, 0, 100))

    slope = model.coef_[0]
    trend = "improving" if slope > 0.5 else ("declining" if slope < -0.5 else "stable")

    return {
        "predicted_next_composite": round(predicted, 2),
        "trend": trend,
        "slope": round(float(slope), 3),
        "at_risk": predicted < FRUSTRATION_COMPOSITE_THRESHOLD,
    }


def composite_score(word_recognition_accuracy, comprehension_avg):
    return round(0.4 * word_recognition_accuracy + 0.6 * comprehension_avg, 2)
