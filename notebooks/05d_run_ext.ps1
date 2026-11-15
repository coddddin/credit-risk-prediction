$env:PYTHONIOENCODING = 'utf-8'
Wait-Process -Id 38936 -ErrorAction SilentlyContinue
Set-Location 'C:\Users\heeju\Desktop\ALGORYITHM\credit_project\notebooks'
& 'C:\Users\heeju\Desktop\ALGORYITHM\credit_project\notebooks\..\.venv\Scripts\python.exe' 05d_candidate_tuning.py AdaBoost+ KNN+ *> 'C:\Users\heeju\Desktop\ALGORYITHM\credit_project\notebooks\05d_candidate_tuning_ext.log'
