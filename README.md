# Streamlit Attendance App

A simple attendance web app built with Streamlit.

## Files

- `streamlit_attendance.py` - main web app
- `student list.xlsx` - student list file to use in the app
- `requirements.txt` - Python packages for Streamlit Cloud

## Run locally

```powershell
streamlit run streamlit_attendance.py
```

## Deploy on Streamlit Cloud

1. Create a new GitHub repository.
2. Upload these files:
   - `streamlit_attendance.py`
   - `requirements.txt`
   - `student list.xlsx`
3. Go to [Streamlit Cloud](https://share.streamlit.io/).
4. Sign in with GitHub.
5. Click `New app`.
6. Select your repository.
7. Set the main file path to `streamlit_attendance.py`.
8. Click `Deploy`.

## Notes

- If `student list.xlsx` is not in the repository, you can still upload it from the sidebar inside the app.
- Attendance can be downloaded as a CSV report.
