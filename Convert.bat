for /R %%I in (*.csv) do (
    csv2xlsx.exe -f %%I
)
pause
