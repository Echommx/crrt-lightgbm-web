# Dynamic Prediction of Antimicrobial Resistance Risk at Repeated Culture Events in Critically Ill Patients Receiving Continuous Renal Replacement Therapy

## Study overview

This repository provides a locally hosted interface for the **LightGBM trend model** developed for the study above. It combines baseline and culture-event characteristics with physiological and laboratory observations from the **24 hours preceding the prediction time (T0)**. The interface reports a study-defined classification of **High risk** or **Low risk**.

## How to use the web application

**1. Install the dependencies and launch the application.**

```bash
python -m pip install -r requirements-inference.txt
python lightgbm_only_web.py --no-browser
```

Open **http://127.0.0.1:8501** in your browser.

**2. Prepare the input.** Click **Download Excel example** to obtain the simulated `.xlsx` workbook. Keep the `trend_input` worksheet and existing column names. Enter the relevant patient and culture-event characteristics, with **one row per observation** for measurements collected within the 24 hours before T0. Do not include information obtained after T0.

**3. Generate a prediction.** Click **Upload Excel (.xlsx)**, select the completed workbook, and choose **Predict**. The result is displayed as **High risk** or **Low risk**, using the configured research threshold.

The on-screen form supports quick entry of the latest measurements. The Excel workflow is preferred for trend-based predictions because it preserves longitudinal observations.

## Interpretation and future updates

The displayed risk category is a **research classification**, not confirmation of antimicrobial resistance and not a substitute for culture, susceptibility testing, or clinical judgment. The example workbook contains **synthetic data only**; no original patient-level research datasets are distributed. Clinical inputs should be appropriately authorized and de-identified.

Additional high-quality samples may support subsequent refinement and assessment of model transportability. Any updated model should undergo standardized outcome review, evaluation of discrimination and calibration, and independent validation before deployment. Increasing sample size alone does not ensure better performance.

**Intended use:** Research and demonstration only.
