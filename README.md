# BrightScript-Json-Beautifier
A small utility function to format JSON strings for Roku

##Function Signature:
```
Function BeautifyJsonString(jsonString As String, indentSpacesNumber = 4) As String
```

## Warning
The function works slow on big JSON strings. For example, JSON string of **~33000 characters** is beautified within **3 seconds**, whereas JSON string of **~116000 characters** is processed within **40 seconds**.

Due to single-threaded Roku's environment, **parsing long JSON strings will hang your UI**.
There is, however, a potential for boosting performance. Please write me to [nikolay.mamaev@gmail.com](mailto:nikolay.mamaev@gmail.com) if you are interested in this improvement.

Thank you!
