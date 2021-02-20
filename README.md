# CODE REFACTOR EFFECT ON PERFORMANCE SPEED OF ANALYSIS

---
## Refactored VBA code to determine what impact the new single loop code had on overall processing time.
---

We were challenged to update the VBA code for the stock analysis to determine if the new refractored code, which was designed to loop through the data only once, would run faster then the original code written with loops.  
---
Original Code

![Original Code](https://user-images.githubusercontent.com/71041680/108607805-a1162500-7390-11eb-8ab0-d3fb4c514a34.PNG)
---
Refactored Code

![Refactored Code](https://user-images.githubusercontent.com/71041680/108607824-bc813000-7390-11eb-8f32-f9ff0987afc4.PNG)


The following images exhibit the refactored code run time, which was significantly faster then the original code.  Run times for the Original VBA code were 0.664 seconds and 0.750 seconds for 2017 + 2018 respectively.  Upon refactoring the code, the new run times were 0.086 and 0.094 for 2017 + 2018.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/71041680/108610364-29052a80-73a3-11eb-9fcf-89b23cca22f9.png)
---
![VBA_Challenge_2018](https://user-images.githubusercontent.com/71041680/108610371-31f5fc00-73a3-11eb-821d-397771ed1c3e.png)


### Advantages and Disadvantages of refactoring code:  

Our analysis for this project has determined that a single loop code will be processed significantly faster than a code with several loops.  Programming the code to run through the data one for only one cycle clearly is more efficient in this instance, as it takes fewer steps and uses less memory.

However, 

### Pros and Cons of refactoring the orignial script:
Pros

Cons






