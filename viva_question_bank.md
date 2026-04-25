# Viva Question Bank With Simple Answers

## 1. What is your project about?
Our project is about improving productivity in a manufacturing line by using lean manufacturing and simulation.  
We used real downtime and production data from a beverage bottling line to find the best improvement method.

## 2. Why did you choose this topic?
We chose this topic because many industries face productivity loss due to stoppages and inefficiency.  
We wanted to study how data and simulation can help select the right lean tool.

## 3. What is the main objective of your project?
The main objective was to identify the biggest source of productivity loss and test lean techniques like 5S and TPM to improve output.

## 4. What is lean manufacturing?
Lean manufacturing is a method used to reduce waste and improve process flow.  
Its aim is to increase value and reduce non-productive activities.

## 5. Why did you use simulation?
We used simulation to test improvements safely before applying them in a real factory.  
It saves time, cost, and risk.

## 6. Why did you use a beverage bottling dataset?
Because it is a real industrial dataset with production and downtime records.  
It helped us do a practical and data-based study instead of a purely theoretical one.

## 7. What data did you use?
We used daily production summaries, hourly production records, and downtime-event logs.  
The dataset covered 56 production days and 1,388 downtime events.

## 8. What are the main performance measures in your project?
The main measures are production output, operating time, pause time, availability, and productivity improvement.

## 9. What is availability?
Availability is the ratio of actual operating time to total monitored time.  
It tells us how much of the total time the system was really working.

## 10. What is downtime?
Downtime is the time when the machine or process is not producing.  
It directly reduces productivity.

## 11. How did you classify downtime?
We divided downtime into three groups: micro-stoppages, minor stoppages, and major stoppages.  
This helped us understand which stoppages were frequent and which caused the biggest loss.

## 12. What is the difference between micro and major stoppages?
Micro-stoppages are short interruptions that happen often.  
Major stoppages are less frequent but usually cause much higher production loss.

## 13. What did you find from the downtime analysis?
We found that major stoppages caused most of the total downtime loss even though they happened less often.  
That was the key reason TPM gave a stronger result.

## 14. What is 5S?
5S is a lean tool used for workplace organization and cleanliness.  
It helps reduce small delays, confusion, and minor interruptions.

## 15. What is TPM?
TPM stands for Total Productive Maintenance.  
It focuses on reducing machine breakdowns and improving equipment reliability.

## 16. Why did TPM perform better than 5S?
TPM performed better because major stoppages were causing most of the total downtime.  
Since TPM targets machine-related losses, it had a bigger impact on productivity.

## 17. What is the integrated lean bundle?
It is a combined improvement approach where we reduce micro, minor, and major stoppages together.  
That is why it gave the highest improvement.

## 18. What algorithm did you use in this project?
We used a bootstrap simulation algorithm based on real production-day profiles.  
It repeatedly sampled actual days from the dataset and recalculated output after applying lean-improvement rules.

## 19. What is the logic of the algorithm?
The logic is simple: first read real production and downtime data, then classify stoppages, then apply lean-based reductions to those stoppages, and finally recalculate production and efficiency.  
This helps us estimate what improvement may happen before actual implementation.

## 20. How did you use lean manufacturing on the data?
We linked each lean tool to a specific type of downtime seen in the dataset.  
For example, 5S was applied to small stoppages, TPM was applied to major stoppages, and the integrated lean approach was applied to all stoppage types together.

## 21. What was the baseline system performance?
The system produced 212,358 liters with an overall availability of 42.66%.  
This means the line had a lot of lost productive time.

## 22. What were the final simulation results?
5S improved output by 2.33%, TPM improved output by 20.61%, and the integrated lean bundle improved output by 32.19%.

## 23. What is the main conclusion of your project?
The main conclusion is that lean tools should be selected based on actual data.  
For this case, TPM should be given first priority because major stoppages were the biggest problem.

## 24. What are the practical uses of this project?
This project can help industries decide which lean tool to apply first.  
It can save time, reduce unnecessary trials, and improve production planning.

## 25. What are the limitations of your project?
The main limitation is that the improvement values were simulated assumptions and not actual implementation results.  
Also, the dataset did not include detailed root-cause information for every stoppage.

## 26. What can be done in future work?
Future work can include root-cause analysis, quality data, labor data, and actual implementation results from the shop floor.

## 27. How is this project useful for manufacturing engineering?
It combines lean manufacturing, productivity analysis, downtime study, and simulation, which are all core areas of manufacturing engineering.

## 28. If the teacher asks “Why not directly implement lean instead of simulation?”
Direct implementation in a real factory may be costly and risky.  
Simulation helps us test improvements first and choose the best option before actual implementation.

## 29. If the teacher asks “What is the novelty of your project?”
The novelty is that we used a real industrial dataset and compared lean techniques through simulation.  
So, the project is practical, data-driven, and not only theoretical.

## 30. If the teacher asks “What did you personally do in the project?”
I worked on understanding the data, analyzing downtime, calculating performance metrics, and preparing the simulation-based comparison of lean scenarios.

## 31. If the teacher asks “Why is major stoppage more important than frequent small stoppages?”
Because a small number of major stoppages can consume more total time than many short stoppages.  
In our project, that is exactly what happened.

## 32. If the teacher asks “Can this method be used in other industries?”
Yes, the same approach can be used in any manufacturing system where production and downtime data are available.  
Only the specific lean priorities may change from one industry to another.

## 33. Best Final Answer if Teacher Says “Explain Your Project in Short”
Our project used real production and downtime data from a bottling line to study productivity loss. We applied lean manufacturing concepts and simulation to compare 5S, TPM, and an integrated lean approach. We found that TPM gave much better improvement than 5S alone because major stoppages were causing most of the production loss.
