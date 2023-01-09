# Module 2-stocks-analysis

## Purpose
The purpose of this assignment is to use Excel and VBA to analyze stock data. 
The assignment specifically asked to refactor code already used in the module and see which version of the code ran faster.

## Results
The code in the module(not refactored) ran, on average, about a second slower than the refactored code.
For instance, this was runtime for the two codes in 2017: </br>
![image](https://user-images.githubusercontent.com/33528884/211237133-36da1477-63e2-4b74-ac1d-f7abd67318e0.png)

![VBA_challenge_2017](https://user-images.githubusercontent.com/33528884/211237182-abf7ab53-60dc-45dd-83f5-2f1ed40f04a9.png)
</br>
</br>
This was the runtimes for the code in 2018: </br>
![image](https://user-images.githubusercontent.com/33528884/211237477-b1802283-9726-4608-bd48-6938b7b5973c.png)

![VBA_challenge_2018](https://user-images.githubusercontent.com/33528884/211237486-e5dc8214-217b-462c-87cb-cb48683bd700.png)

## Summary
Refactoring code means to improve the overall structure of the code without fundamentally altering its overall job. For this reason, one obvious benefit of refactoring is that, at the end of the day, it is still the same code that has been worked on in the past. Another benefit is that refactoring does not require developers to maintain a copy as the refactoring happens on the original code. Refactoring also does not have to be done on the full code. If a developer is only working on a part of the code, they can choose to refactor just that part. Since the basic idea of refactoring is to change complex code into simpler code, one downside to this is that there will inevitably more code to handle and test. For instance, in this project, one nested for loop was broken down into three seperate for loops. Whilst each for loop is, by itself, simpler, there is also more code to manage now. 
