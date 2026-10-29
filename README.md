# Course recommender system: content-based, clustering and collaborative filtering on 233k enrolments

Capstone of the IBM Machine Learning Professional Certificate (October 2026). The task: recommend courses to the learners of an online catalogue, compare several recommender families, and present the results as a 26-slide report. Along the way the project found and fixed three bugs in the course's reference implementations and showed that the supplied rating file carries no learnable signal.

| | |
|---|---|
| Report (26 slides) | [`report/Machine Learning Capstone Project Report.pdf`](report/Machine%20Learning%20Capstone%20Project%20Report.pdf) (also as `.pptx`) |
| Benchmark notebook | [`notebooks/benchmark_collaborative_filtering.ipynb`](notebooks/benchmark_collaborative_filtering.ipynb): every collaborative-filtering model on one common split |
| Results | `results/` (every number in the report as CSV) |
| Data | `data/course_genre.csv`, `data/course_processed.csv` (course catalogue with 14 genre flags and course text) |

The eleven guided lab notebooks that built each recommender step by step are IBM course templates and are not republished here; the benchmark notebook, the report, its generator and all result tables are this project's own work.

## The data

307 courses tagged with 14 genres, 33,901 learners, 233,306 enrolments. The catalogue is dominated by back-end development (78 courses), machine learning (69) and databases (60), and the top 20 courses account for 63% of all enrolments, which is why popularity-blind recommenders matter here.

## Recommenders built

| Family | Approach | Output on the 33,901 test users |
|---|---|---|
| Content-based, user profile | genre-weighted user profile × course genre vectors, score threshold 20 | 14.1 recommendations per user, 16,554 users reached, 280 of 307 courses covered |
| Content-based, course similarity | bag-of-words cosine similarity between course descriptions | similar-course lists, e.g. `results/similar_to_ML0101ENv3.csv` |
| Clustering | PCA (9 components) then K-means (k = 8) on user profiles; recommend what a user's cluster takes | 13.7 per user, 33,381 users reached, but only 50 distinct courses (popularity threshold 500) |
| Collaborative filtering | KNN (Surprise), NMF, SVD, two neural embedding models, ridge regression on embeddings | see the benchmark below |
| Rating-mode classification | predict a user's typical rating from user and course embeddings with 6 classifiers | accuracy 33–34% = chance, see below |

## Collaborative filtering benchmark (RMSE on one common test split)

| Model | `ratings.csv` (3/4/5 scale) | lab-05 ratings (2/3 scale) |
|---|---|---|
| Global-mean baseline | 0.814 | 0.210 |
| KNN item-based, cosine, k = 40 | 0.892 | 0.193 |
| NMF, 16 factors | 0.945 | |
| SVD | 0.839 | |
| RecommenderNet, as in the lab | 1.289 | 0.119 |
| **RecommenderNetV2** (per-example dot product + dense head) | 0.814 | **0.076** |
| Ridge regression on the learned embeddings | 0.814 | |

Two findings follow from the table:

1. **On `ratings.csv` nothing beats the global mean.** The file has the same 233,306 (user, course) pairs as the lab's 2/3-scale file but its 3/4/5 ratings are uncorrelated with them (r = −0.003); the rating-mode classifiers landing at chance (33%) confirm it. The report says so rather than tuning models against noise.
2. **The reference neural model had a bug.** `tf.tensordot(user, item, 2)` sums the dot products over the whole batch, so every example in a batch gets the same score. RecommenderNetV2 computes a per-example dot product and adds a dense head; on the lab ratings it cuts RMSE from 0.119 to 0.076 in three epochs.

Two more fixes in the labs: the Surprise `rating_scale` was set to (2, 3) where the data was (3, 5), and NMF's suggested `init_high = 5` starts the factor products above the scale so every prediction clips to 5 (RMSE 1.24 before the fix, 0.95 after).

## Reproduce

```bash
pip install -r requirements.txt
jupyter notebook notebooks/benchmark_collaborative_filtering.ipynb
```

`report/source/build_report.py` builds the slide deck with python-pptx from `results/*.csv` (and a word cloud of the course titles); `export_pdf.ps1` exports it to PDF with PowerPoint. The rating files and the learned embeddings (about 20 MB) are IBM lab assets and are not in the repo; the benchmark notebook expects them in `../datasets/`.
