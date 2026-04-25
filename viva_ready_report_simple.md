# Viva-Ready Report

## Project Title
**Simulation-Based Implementation of Lean Manufacturing Techniques for Productivity Improvement: A Data-Driven Case Study From a Beverage Bottling Line**

## Submitted By
- Divash Krishnam
- Ashutosh Verma
- Raja

## Guided By
**Dr. MOHD SHUAIB**, Mechanical Department

## 1. Project in Simple Words
This project is about improving production in a manufacturing line by finding where time is being wasted.  
We used real production and downtime data from a beverage bottling line and tested different lean manufacturing techniques through simulation to see which one gives the best improvement.

## 2. Why We Chose This Project
In many industries, machines do not work continuously. They stop again and again due to small issues, delays, or breakdowns.  
Instead of applying lean tools blindly, we wanted to use data and simulation to find which improvement method would actually give the highest productivity gain.

## 3. Main Objective
The main objective of the project was to identify the biggest causes of productivity loss and then test lean methods like 5S, TPM, and an integrated lean approach to see which one improves output the most.

## 4. Problem Statement
The line was producing less than its possible capacity because a lot of time was lost in stoppages.  
The main problem was to decide which lean technique should be given first priority for improving productivity.

## 5. Data Used in the Project
We used a real industrial dataset from a beverage bottling line.

### Important details of the dataset
- Monitoring period: July 22, 2022 to February 6, 2023
- Production summary records: 59
- Calendar production days: 56
- Hourly operating observations: 265
- Downtime events: 1,388
- Product sizes: 3 L and 5 L

## 6. What Is Lean Manufacturing in This Project?
Lean manufacturing means reducing waste and improving process flow.  
In this project, the main waste was lost production time due to downtime.

### Lean tools used in our study
- **5S and Visual Management**: To reduce small interruptions and improve workplace organization.
- **TPM (Total Productive Maintenance)**: To reduce major machine stoppages and improve reliability.
- **Integrated Lean Bundle**: A combination of multiple lean actions to reduce both small and large stoppages.

## 7. Why We Used Simulation
If we directly apply a new method in a real factory, it may cost time and money.  
Simulation helps us test changes safely before actual implementation.

### Simple meaning
Simulation means creating a model of the real system and checking what happens if we improve certain parts of it.

## 8. Step-by-Step Methodology
The project was completed in the following steps:

1. We collected and studied the production and downtime data.
2. We calculated important measures like production, operating time, pause time, and efficiency.
3. We divided downtime into three groups:
   - Micro-stoppages: up to 2 minutes
   - Minor stoppages: 2 to 10 minutes
   - Major stoppages: above 10 minutes
4. We checked which type of stoppage caused the highest loss.
5. We created simulation scenarios for 5S, TPM, and integrated lean.
6. We compared the results of all scenarios with the original system.

## 9. Algorithm and Logic Used in the Project
The main algorithm used in this project was a **bootstrap simulation algorithm based on real production data**.

### Simple logic of the algorithm
1. Read the real production and downtime data.
2. Convert downtime into minutes and group it date-wise.
3. Calculate daily production rate, operating time, and pause time.
4. Build a baseline model from the actual observed days.
5. Randomly resample these real day profiles many times.
6. Apply a lean rule to the downtime data:
   - remove some micro-stoppages for 5S,
   - reduce major stoppage duration for TPM,
   - reduce all types of stoppages for integrated lean.
7. Recalculate output and efficiency after each change.
8. Compare the new results with the baseline system.

### One-line explanation for viva
We used a data-driven bootstrap simulation, where actual production days were repeatedly sampled and modified according to lean-improvement rules to estimate productivity improvement.

## 10. How Lean Manufacturing Was Used on the Data
Lean manufacturing was not used only as a theory in this project.  
We directly connected lean tools to the downtime data.

### How each lean tool was linked to the data
- **5S + Visual Management** was linked to micro-stoppages because it helps reduce small avoidable interruptions.
- **TPM** was linked to major stoppages because it improves machine reliability and reduces breakdown duration.
- **Integrated Lean Bundle** was linked to micro, minor, and major stoppages together.

### Simple meaning
We treated downtime as waste in the data, and then used lean techniques as methods to reduce that waste inside the simulation model.

## 11. Important Formulae Used

### Availability / Efficiency
\[
A = \frac{T_{op}}{T_{mon}}
\]
This means:
- `T_op` = operating time
- `T_mon` = total monitored time

### Downtime
\[
T_{dt} = T_{mon} - T_{op}
\]
This tells us how much time was lost.

### Production Rate
\[
R = \frac{Q}{T_{op}}
\]
This tells us how much production happens in one hour of actual machine operation.

### Productivity Improvement
\[
\%\Delta Q = \frac{Q' - Q}{Q}\times 100
\]
This tells us the percentage improvement after applying a lean scenario.

## 12. Main Findings From the Real Data

### Baseline system performance
- Total production: **212,358 liters**
- Total units produced: **60,888**
- Total monitored time: **238.65 hours**
- Total operating time: **101.81 hours**
- Total pause time: **137.56 hours**
- Overall availability: **42.66%**

### What this means
The line was active for the full monitoring period, but actual useful operating time was much lower.  
This clearly shows that downtime was the main reason for low productivity.

## 13. Downtime Analysis
We found that not all stoppages had the same impact.

### Downtime categories
- Micro-stoppages: **851 events**, but only **13.59%** of total downtime minutes
- Minor stoppages: **443 events**, and **20.15%** of total downtime minutes
- Major stoppages: **94 events**, but **66.26%** of total downtime minutes

### Simple conclusion
Small stoppages happened many times, but major stoppages caused the biggest actual loss.  
So, if we only focus on small interruptions, we may not get the highest productivity improvement.

## 14. Simulation Scenarios

### Scenario 1: 5S + Visual Management
We assumed that better organization and control would remove some small stoppages.

### Scenario 2: TPM
We assumed that better maintenance would reduce the duration of major breakdowns.

### Scenario 3: Integrated Lean Bundle
We assumed that a combination of improvements would reduce micro, minor, and major stoppages together.

## 15. Simulation Results

| Scenario | Output (L) | Improvement | Efficiency |
|---|---:|---:|---:|
| Baseline | 211,924.69 | 0.00% | 42.37% |
| 5S + Visual Management | 216,866.69 | 2.33% | 43.25% |
| TPM | 255,604.10 | 20.61% | 50.81% |
| Integrated Lean Bundle | 280,142.22 | 32.19% | 55.33% |

## 16. Interpretation of Results
The most important result is that TPM gave much better improvement than 5S alone.  
This happened because major stoppages were causing most of the lost time, and TPM directly targets major stoppages.

The integrated lean bundle gave the best overall result because it reduced losses in all categories.

## 17. Final Conclusion
The project shows that lean tools should not be selected only by habit or general opinion.  
They should be selected based on the actual loss pattern in the data.

For this bottling line:
- 5S was helpful, but its effect was small.
- TPM was much more powerful because major stoppages were the real problem.
- The integrated lean approach gave the highest total improvement.

So, the best recommendation is:
1. First focus on TPM and machine reliability.
2. Then use 5S and visual management.
3. Finally combine both for larger system-wide improvement.

## 18. What We Learned From This Project
- How to use real industrial data for manufacturing analysis
- How to calculate productivity and downtime metrics
- How to apply lean manufacturing concepts in a practical way
- How simulation can help in decision-making
- How to connect data analysis with process improvement

## 19. Limitations of the Project
- The improvement values were simulated assumptions, not actual factory implementation results.
- The dataset did not tell the exact root cause of every stoppage.
- The study was based on one industrial case, so results may vary in another factory.

## 20. Future Scope
- Add root-cause analysis of downtime
- Include quality defects and labor data
- Test the same method in other manufacturing industries
- Validate the results with real shop-floor implementation

## 21. 30-Second Viva Summary
Our project used real production and downtime data from a beverage bottling line to find where productivity was being lost.  
We tested lean methods through simulation and found that TPM gave much better improvement than 5S alone because major stoppages were causing most of the downtime loss.

## 22. 1-Minute Viva Explanation
This project is based on lean manufacturing and simulation. We used a real industrial dataset from a beverage bottling line and studied production, downtime, and efficiency. Then we classified downtime into micro, minor, and major stoppages. After that, we simulated three lean scenarios: 5S, TPM, and an integrated lean bundle. The main result was that major stoppages caused most of the production loss, so TPM gave much higher productivity improvement than 5S alone. This shows that data-based lean implementation is more effective than applying tools without proper analysis.

## 23. If Teacher Asks “Why Is This Project Important?”
This project is important because many industries use lean tools without checking which problem is actually the biggest.  
Our project shows how to use real data and simulation to choose the right improvement method scientifically.

## 24. If Teacher Asks “What Is the Main Contribution?”
The main contribution is that we used real industrial data and simulation together to compare lean techniques.  
Instead of only discussing lean theoretically, we showed which method gives the best practical improvement.
