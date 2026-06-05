# Carbon-Aware Machine Learning Model Selection for Germany

This project package was generated for the PS26 final project prompt.

## Core Files

- `data.csv`: raw OPSD hourly power-system time-series file.
- `preprocessed_content.csv`: Germany-focused engineered modeling table.
- `notebooks/`: five runnable notebooks, one per research question.
- `outputs/tables/`: CSV result tables and project metadata.
- `outputs/figures/`: publication figures as PDF plus PNG companions for documents and posters.
- `icons/`: reusable SVG icons used in the workflow figure.
- `icons/google_slides_png/`: high-resolution PNG icon exports for slides.
- `posters/carbon_aware_ml_academic_posters_v2.pptx`: editable poster source deck.
- `posters/carbon_aware_ml_poster_final.pdf`: poster PDF for final assignment upload.
- `presentations/carbon_aware_ml_graphical_abstract_google_slides.pptx`: editable Google-Slides-ready deck.
- `presentations/copyable_svg_icon_bank_google_slides.pptx`: icon bank for quick reuse in presentation layouts.
- `final_submission_links.txt`: ready-to-submit list of required links.
- `report/Carbon_Aware_ML_Germany_IEEEPES_Technical_Report.docx`: IEEE/PES-style technical report.

## Research Questions

- **RQ1**: Which model family gives the strongest Pareto trade-off between predictive accuracy and estimated computational carbon for German next-day high-load-event classification?
- **RQ2**: How much do load history, renewable generation, TSO-zone state, and day-ahead forecast features improve performance relative to a calendar-only baseline?
- **RQ3**: Are the model rankings stable across German winter, spring, summer, and autumn operating regimes?
- **RQ4**: Can an interpretable low-complexity model explain the dominant drivers of high-load risk without a large loss in predictive utility?
- **RQ5**: How much carbon-proxy reduction is available if model training is scheduled into low-residual-load windows rather than arbitrary or high-residual-load hours?

## Reproduce

Run:

```bash
python3 scripts/generate_analysis_assets.py
```

The notebooks call the same generation functions and can be run independently.

## Submission Checklist

This repository is organized to satisfy the required GitHub submission format:

- Executed notebook code: `notebooks/rq1_carbon_aware_ml.ipynb` through `notebooks/rq5_carbon_aware_ml.ipynb`
- Displayed notebook outputs: each RQ notebook includes the result table and figure output inline for GitHub preview
- Saved PDF figures: `outputs/figures/*.pdf`
- Saved CSV tables: `outputs/tables/*.csv`
- Reproducible asset script: `scripts/generate_analysis_assets.py`
