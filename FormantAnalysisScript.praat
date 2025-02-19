Read from file: "C:\ProgrammeOhneInstallation\tempraatfile.wav"
To Formant (burg): 0, 5, 5500, 0.025, 50
formant1f$ = Get mean: 1, 0, 0, "hertz"
formant2f$ = Get mean: 2, 0, 0, "hertz"
formant3f$ = Get mean: 3, 0, 0, "hertz"
formant4f$ = Get mean: 4, 0, 0, "hertz"
formant5f$ = Get mean: 5, 0, 0, "hertz"
formant1b$ = Get quantile of bandwidth: 1, 0, 0, "hertz", 0.5
formant2b$ = Get quantile of bandwidth: 2, 0, 0, "hertz", 0.5
formant3b$ = Get quantile of bandwidth: 3, 0, 0, "hertz", 0.5
formant4b$ = Get quantile of bandwidth: 4, 0, 0, "hertz", 0.5
formant5b$ = Get quantile of bandwidth: 5, 0, 0, "hertz", 0.5
formantlist$ = List: "no", "yes", 3, "no", 1, "yes", 1, "yes"
formants$ = formant1f$ - " hertz" + ";" + formant2f$ - " hertz" + ";"
... + formant3f$ - " hertz" + ";" + formant4f$ - " hertz" + ";"
... + formant5f$ - " hertz" + ";" + formant1b$ - " hertz" + ";" + formant2b$ - " hertz" + ";"
... + formant3b$ - " hertz" + ";" + formant4b$ - " hertz" + ";"
... + formant5b$ - " hertz"
writeInfoLine: formants$
# writeInfoLine: formantlist$
