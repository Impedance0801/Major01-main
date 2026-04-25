# Research Paper Topic Options and Draft

## 1. Topic Options You Can Work On

1. **Data-Driven Simulation of Lean Interventions for Productivity Improvement in SME Bottling Lines**
   Research gap: Most lean-simulation studies in SMEs use proprietary or hypothetical data; few use open, event-level industrial datasets with downtime logs.

2. **Digital Value Stream Mapping for Small-Scale Manufacturing Using IIoT Production Data**
   Research gap: Value stream mapping is still often static and manually updated; recent digital twin work is promising, but SME-oriented empirical validations remain limited.

3. **TPM-Driven Downtime Reduction in Small Manufacturing Systems: A Simulation-Based Evaluation**
   Research gap: Many studies discuss TPM conceptually, but fewer quantify how reducing long breakdowns changes output under real operational variability.

4. **5S, Standardized Work, and Micro-Stoppage Reduction: An Event-Log-Based Study**
   Research gap: Micro-stoppages are frequent but under-modeled in lean studies, especially in food and beverage and other SME-like operations.

5. **Integrated Lean and Smart Manufacturing for Small Enterprises: A Reproducible Open-Data Framework**
   Research gap: The lean-smart manufacturing literature is growing, but there is still a shortage of reproducible, open-data case studies that small firms can adapt directly.

## 2. Recommended Topic

Your original idea is strong. A sharper, more researchable version is:

**Simulation-Based Implementation of Lean Manufacturing Techniques for Productivity Improvement: A Data-Driven Case Study Using a Real Beverage Bottling Dataset**

This keeps your original theme, but makes it publishable because it adds:

- a real industrial dataset,
- a reproducible simulation method,
- explicit research gap coverage,
- measurable lean scenarios.

## 3. Research Gap

Based on the literature, the main gap is:

**Existing studies show that simulation helps prioritize lean tools, but most studies are single-case, proprietary, static, or not reproducible. Very few studies combine real event-level production data, downtime logs, and simulation-based testing of lean interventions in SME-relevant manufacturing settings.**

More specifically:

1. SME-focused simulation literature still reports a need for solutions targeted to SMEs rather than large, resource-rich firms [3].
2. Lean-simulation studies often evaluate tool combinations, but many rely on closed industrial data that other researchers cannot reuse [4], [5].
3. New work on digital twins and smart manufacturing shows potential for real-time lean support, but empirical integration with downtime-driven lean simulation is still limited [6], [10].
4. Food and beverage SME research shows lean relevance, yet quantitative simulation studies in this sector remain relatively scarce [9].

## 4. Draft Research Paper

## Title

**Simulation-Based Implementation of Lean Manufacturing Techniques for Productivity Improvement: A Data-Driven Case Study from a Beverage Bottling Line**

## Abstract

Lean manufacturing is widely recognized as an effective approach for reducing waste and improving productivity, yet many small and medium-sized manufacturing firms struggle to identify which lean techniques will deliver the greatest benefit before implementation. This study develops a data-driven simulation framework to evaluate lean interventions using a real industrial production dataset from a beverage bottling line monitored through an Industrial Internet of Things (IIoT) system. The dataset contains 59 production records across 56 calendar days between July 2022 and February 2023, 265 hourly operating observations, and 1,388 downtime events. First, the operational profile of the line is analyzed through production, efficiency, and downtime metrics. Next, an event-level bootstrap simulation is used to test three lean scenarios: 5S with visual management, total productive maintenance (TPM), and an integrated lean bundle combining micro-stoppage reduction, minor-stop reduction, and major-stop reduction. The baseline system produced 212,358 liters with an overall availability of 42.66%. Downtime analysis showed that micro-stoppages represented 61.31% of all events but only 13.59% of downtime minutes, while major stoppages represented only 6.77% of events but 66.26% of downtime minutes. Simulation results indicate that 5S and visual management yield a modest output gain of 2.33%, TPM yields 20.61%, and the integrated lean bundle yields 32.19%, while increasing simulated availability from 42.37% to 55.33%. The results show that, for this line, reducing major downtime events produces much greater productivity gains than focusing only on frequent short interruptions. The study contributes a reproducible open-data approach for simulation-based lean evaluation and offers a practical roadmap for SME-oriented manufacturing improvement.

## Keywords

Lean manufacturing; discrete-event simulation; productivity improvement; small and medium enterprises; bottling line; downtime analysis; total productive maintenance; 5S; digital manufacturing; IIoT

## 1. Introduction

Manufacturing firms continue to adopt lean practices to improve throughput, reduce waste, and increase responsiveness. However, implementation is often difficult in small and medium-sized enterprises (SMEs) because they typically face limited capital, limited analytical capability, and limited tolerance for failed experiments. For that reason, simulation is valuable: it allows managers to test lean alternatives before disrupting actual production [3], [5].

Recent literature confirms that lean tools can improve operational performance, but their impact is context dependent [1], [4], [5]. Some studies report that simulation can identify which tools are most effective in a particular setting [5]. Others show that lean and smart manufacturing increasingly reinforce each other, especially when digital data are available [6], [10]. At the same time, the SME simulation literature still emphasizes the lack of SME-oriented solutions and reproducible empirical cases [3].

This study addresses that gap by using a real open industrial dataset from a beverage filling line collected through an IIoT monitoring system [7], [8]. Unlike many prior studies, the present work is built on publicly reusable event-level downtime data and transparent scenario assumptions. The objective is to evaluate how different lean interventions may improve productivity under real operational variability.

The specific objectives are:

1. To characterize the production and downtime behavior of the bottling line.
2. To identify the dominant operational waste pattern.
3. To simulate the impact of selected lean interventions on productivity and availability.
4. To propose a reproducible framework suitable for SME-oriented manufacturing research.

## 2. Literature Review

Lean manufacturing has been shown to improve operational performance across multiple manufacturing contexts [1]. However, the effect of lean tools is not uniform, and firms often struggle to know which tools should be prioritized first. Possik et al. [4] showed through co-simulation that some tools such as 5S may be broadly useful, while others such as SMED or cross-training are more context dependent. Similarly, Frecassetti et al. [5] demonstrated in an Italian SME that simulation can be used to identify the combination of lean practices that provides the strongest process gains.

Simulation is especially relevant to SMEs because it supports decision-making without requiring risky physical trials. Yet Yu and Chen [3] concluded that simulation solutions for SMEs still face important adoption barriers and that more targeted solutions are needed. Oleghe and Salonitis [2] further argued that quantitative lean modeling remains underdeveloped, particularly when human and process dynamics interact.

The literature also shows a growing convergence between lean and smart manufacturing. Bokhorst et al. [6] found that smart manufacturing is frequently built on lean principles rather than replacing them. Recent digital twin work on automated value stream mapping also suggests that real-time data can strengthen lean diagnosis and improvement selection [10].

Despite these advances, an important gap remains. Most lean-simulation studies rely on closed case-company data. Reproducible open-data studies are rare, especially in food and beverage and SME-relevant settings. This is important because SMEs need simple, transferable, and evidence-based improvement frameworks. The present study therefore combines an open industrial dataset [8] with a simulation-based lean analysis to produce a transparent and reusable case.

## 3. Key Mathematical Formulae

The main performance relationships used in the study are:

### 3.1 Availability / Operational Efficiency

\[
A = \frac{T_{op}}{T_{mon}}
\]

Where:

- \(A\) = operational availability or efficiency
- \(T_{op}\) = effective operating time
- \(T_{mon}\) = monitored time

### 3.2 Downtime

\[
T_{dt} = T_{mon} - T_{op}
\]

Where \(T_{dt}\) is downtime or pause time.

### 3.3 Production Rate

\[
R = \frac{Q}{T_{op}}
\]

Where:

- \(R\) = production rate in liters per operating hour
- \(Q\) = liters produced

### 3.4 Simulated Output

\[
Q' = R \times (T_{mon} - T'_{dt})
\]

Where \(Q'\) is simulated output under an improvement scenario.

### 3.5 Scenario-Based Downtime Transformation

\[
T'_{dt} = T_{res} + \sum_{i=1}^{n}\alpha_i t_i
\]

Where:

- \(T_{res}\) = residual downtime not represented at event level
- \(t_i\) = duration of downtime event \(i\)
- \(\alpha_i\) = improvement factor applied to downtime category \(i\)

### 3.6 Productivity Improvement

\[
\%\Delta Q = \frac{Q' - Q}{Q} \times 100
\]

## 4. Research Methodology

### 4.1 Research Design

This study uses a data-driven simulation-based case study design. The analysis combines descriptive statistics, downtime categorization, and bootstrap-based event simulation. The overall logic is:

1. Analyze the real production dataset.
2. Identify dominant waste categories.
3. Define lean interventions matched to those waste categories.
4. Simulate counterfactual production outcomes.
5. Compare output and efficiency against the baseline.

### 4.2 Dataset

The study uses the open dataset **Industrial Production Time-Series Dataset from a Beverage Bottling Line** published on Zenodo in January 2026 [8]. The data were collected from a real industrial filling machine using an IIoT monitoring system, with sensor-based production counts and downtime logging [8].

The dataset contains:

- 59 production summary records,
- 56 calendar production days,
- 265 hourly operating observations,
- 1,388 downtime events,
- monitoring period from **July 22, 2022 to February 6, 2023**.

### 4.3 Variables Used

The following variables were used:

- liters produced,
- production units,
- monitored time,
- operating time,
- pause time,
- hourly efficiency,
- downtime event duration,
- product type (3 L and 5 L).

### 4.4 Baseline Analysis

The baseline system was first characterized using descriptive statistics. Daily and hourly efficiency were computed, monthly trends were visualized, and downtime events were grouped into three categories:

- **Micro-stoppages:** \(t \leq 2\) minutes
- **Minor stoppages:** \(2 < t \leq 10\) minutes
- **Major stoppages:** \(t > 10\) minutes

This classification was selected because it separates frequent short interruptions from low-frequency, high-impact breakdowns.

### 4.5 Simulation Logic

An event-level bootstrap simulation with 5,000 replications was developed. Each replication resampled observed production days with replacement, preserving the real variability of monitored time, production rate, and downtime structure. Event durations were normalized to the recorded daily pause time so that the simulated baseline matched the plant record.

Three lean scenarios were tested:

1. **Scenario 1: 5S + Visual Management**
   Assumption: 15% of micro-stoppages are eliminated through better workplace organization and visual control.

2. **Scenario 2: TPM**
   Assumption: major stoppage durations are reduced by 25% through maintenance discipline and faster recovery.

3. **Scenario 3: Integrated Lean Bundle**
   Assumption: 20% of micro-stoppages are eliminated, minor stoppages are reduced by 15%, and major stoppages are reduced by 30%.

These are literature-informed scenario assumptions rather than observed post-implementation measurements. The purpose is to compare intervention leverage, not to claim guaranteed plant outcomes.

## 5. Results and Graphical Analysis

### 5.1 Baseline Operational Results

The line produced **212,358 liters** from **60,888 units** during **238.65 monitored hours**. Effective operating time was only **101.81 hours**, while downtime was **137.56 hours**, giving an overall availability of **42.66%**. The mean daily efficiency recorded in the summary sheet was **46.52%**.

The 3 L product had an average production rate of **1,963.71 L/h**, while the 5 L product averaged **2,918.68 L/h**. At the date-aggregated level, pause time had a strong negative correlation with efficiency (\(r = -0.707\)), confirming that downtime was a dominant productivity driver.

### 5.2 Downtime Structure

Downtime events were highly skewed:

- Median downtime event: **1.50 minutes**
- 90th percentile: **6.50 minutes**
- 95th percentile: **13.74 minutes**
- Maximum event: **349.83 minutes**

The downtime pattern shows why lean prioritization matters:

- Micro-stoppages: **851 events** and **13.59%** of total downtime minutes
- Minor stoppages: **443 events** and **20.15%** of total downtime minutes
- Major stoppages: **94 events** and **66.26%** of total downtime minutes

This means frequent short stoppages create operational noise, but major stoppages dominate lost time.

### 5.3 Monthly Efficiency Pattern

Monthly average efficiency fluctuated substantially. Lower-performing periods occurred in **August 2022** and **December 2022**, while **January 2023** showed the strongest average performance. This variation supports the need for simulation because the line does not behave as a stable deterministic system.

![Monthly efficiency](C:/codex/researchPaper/Project_01/outputs/monthly_efficiency.png)

**Figure 1.** Monthly average operational efficiency from the real bottling-line dataset.

### 5.4 Downtime Event Distribution

The downtime distribution is strongly right-skewed, with many short interruptions and a small number of extremely long stoppages.

![Downtime distribution](C:/codex/researchPaper/Project_01/outputs/downtime_distribution.png)

**Figure 2.** Distribution of downtime event duration (truncated at 60 minutes for readability).

### 5.5 Simulation Results

The simulation results are summarized in Table 1.

| Scenario | Mean simulated output (L) | Output gain (%) | Mean efficiency (%) | Efficiency gain (pp) | Pause reduction (%) |
|---|---:|---:|---:|---:|---:|
| Baseline | 211,924.69 | 0.00 | 42.37 | 0.00 | 0.00 |
| 5S + Visual Management | 216,866.69 | 2.33 | 43.25 | 0.88 | 1.52 |
| TPM | 255,604.10 | 20.61 | 50.81 | 8.44 | 14.68 |
| Integrated Lean Bundle | 280,142.22 | 32.19 | 55.33 | 12.96 | 22.51 |

The integrated lean bundle produced the largest gain, increasing mean output by **68,217.53 liters** over the simulated baseline across the 59 sampled production records. However, the key insight is that **TPM alone delivered much larger gains than 5S alone**, because major stoppages consumed most of the total downtime.

![Scenario comparison](C:/codex/researchPaper/Project_01/outputs/scenario_output.png)

**Figure 3.** Mean simulated production output across the baseline and lean scenarios.

## 6. Discussion

The results support three important conclusions.

First, lean priorities should be aligned with the downtime structure rather than with event frequency alone. If managers focus only on the most visible recurring interruptions, they may overinvest in micro-stop reduction while neglecting the small number of long stoppages that drive most lost time. In this case, TPM produced far greater benefits than 5S because major events consumed 66.26% of downtime time.

Second, 5S and visual management still have value, but their effect is more incremental. The simulation suggests that workplace organization and visual control can improve flow stability, yet by themselves they do not address the dominant source of productivity loss. This aligns with prior work showing that some lean tools are broadly useful, but other tools must be selected according to context [4].

Third, the study demonstrates the practical value of combining lean thinking with digital manufacturing data. Recent literature on smart manufacturing and digital value stream mapping argues that real-time data can improve lean decision-making [6], [10]. The present results support that view. Instead of selecting lean tools only through observation or expert judgment, managers can use event-level data to quantify which intervention should be prioritized first.

For SME-oriented industries, this is especially useful. SMEs often cannot afford multiple failed improvement projects. A simulation-first approach reduces implementation risk by screening alternatives before physical deployment. This directly addresses the need for SME-targeted simulation support identified by Yu and Chen [3].

## 7. Conclusion

This study developed a simulation-based framework for evaluating lean manufacturing interventions using a real industrial bottling-line dataset. The baseline analysis showed low availability, high downtime, and strong variability in production performance. Although micro-stoppages were the most frequent events, major stoppages dominated total downtime.

The simulation showed that:

1. 5S and visual management produce a modest gain of **2.33%**.
2. TPM produces a much stronger gain of **20.61%**.
3. An integrated lean bundle increases output by **32.19%** and raises mean efficiency from **42.37%** to **55.33%**.

The main managerial implication is clear: in this bottling-line case, **maintenance-led reduction of major stoppages should be the first priority**, followed by broader lean stabilization practices such as 5S, standard work, and visual management.

The main research contribution is methodological. The study presents a transparent, open-data, reproducible approach to simulation-based lean analysis that final-year students and SME practitioners can adapt to other manufacturing lines.

### Limitations and Future Scope

1. The lean improvements were simulated as literature-informed scenarios rather than measured post-implementation outcomes.
2. The dataset does not classify downtime causes, so the mapping from downtime category to lean tool remains inferential.
3. Future research can extend the model by collecting cause-coded downtime data, adding labor and quality variables, and validating the results with real post-intervention shop-floor implementation.

## References

[1] I. Belekoukias, J. A. Garza-Reyes, and V. Kumar, “The impact of lean methods and tools on the operational performance of manufacturing organisations,” *International Journal of Production Research*, vol. 52, no. 18, pp. 5346-5366, 2014. doi: [10.1080/00207543.2014.903348](https://doi.org/10.1080/00207543.2014.903348)

[2] O. Oleghe and K. Salonitis, “Hybrid simulation modelling of the human-production process interface in lean manufacturing systems,” *International Journal of Lean Six Sigma*, vol. 10, no. 2, pp. 665-690, 2019. doi: [10.1108/IJLSS-01-2018-0004](https://doi.org/10.1108/IJLSS-01-2018-0004)

[3] F. Yu and Z. Chen, “Tools, application areas and challenges of factory simulation in Small and Medium-Sized Enterprises - A Review,” *Procedia CIRP*, vol. 104, pp. 399-404, 2021. doi: [10.1016/j.procir.2021.11.067](https://doi.org/10.1016/j.procir.2021.11.067)

[4] J. Possik, A. Zouggar-Amrani, B. Vallespir, and G. Zacharewicz, “Lean techniques impact evaluation methodology based on a co-simulation framework for manufacturing systems,” *International Journal of Computer Integrated Manufacturing*, vol. 35, no. 1, pp. 91-111, 2022. doi: [10.1080/0951192X.2021.1972468](https://doi.org/10.1080/0951192X.2021.1972468)

[5] S. Frecassetti, B. Kassem, K. Kundu, M. Ferrazzi, and A. Portioli-Staudacher, “Introducing Lean practices through simulation: A case study in an Italian SME,” *Quality Management Journal*, vol. 30, no. 2, pp. 90-104, 2023. doi: [10.1080/10686967.2023.2171326](https://doi.org/10.1080/10686967.2023.2171326)

[6] J. A. C. Bokhorst, W. Knol, J. Slomp, and T. Bortolotti, “Assessing to what extent smart manufacturing builds on lean principles,” *International Journal of Production Economics*, vol. 253, 108599, 2022. doi: [10.1016/j.ijpe.2022.108599](https://doi.org/10.1016/j.ijpe.2022.108599)

[7] G. A. David, P. M. C. Monson, C. Soares Junior, P. O. Conceição Junior, P. R. Aguiar, and A. Simeone, “IoT-Driven Deep Learning for Enhanced Industrial Production Forecasting,” *IEEE Internet of Things Journal*, vol. 11, no. 23, pp. 38486-38495, 2024. doi: [10.1109/JIOT.2024.3447579](https://doi.org/10.1109/JIOT.2024.3447579)

[8] G. A. David, P. M. C. Monson, and C. Soares Junior, *Industrial Production Time-Series Dataset from a Beverage Bottling Line*, Zenodo, 2026. doi: [10.5281/zenodo.18146866](https://doi.org/10.5281/zenodo.18146866)

[9] J. M. Matindana and M. J. Shoshiwa, “Lean manufacturing implementation in food and beverage SMEs in Tanzania: using structural equation modelling (SEM),” *Management System Engineering*, vol. 4, no. 1, pp. 1-14, 2025. doi: [10.1007/s44176-025-00036-3](https://doi.org/10.1007/s44176-025-00036-3)

[10] M. Reslan, M. J. Triebe, R. Venketesh, and A. J. Hartwell, “Automation of Value Stream Mapping: A Case Study on Enhancing Lean Manufacturing Tools Through Digital Twins,” *Procedia CIRP*, vol. 134, pp. 455-460, 2025. doi: [10.1016/j.procir.2025.02.159](https://doi.org/10.1016/j.procir.2025.02.159)

[11] C. Unal and S. Bilget, “Examination of lean manufacturing systems by simulation technique in apparel industry,” *The Journal of The Textile Institute*, vol. 112, no. 3, pp. 377-387, 2021. doi: [10.1080/00405000.2020.1756104](https://doi.org/10.1080/00405000.2020.1756104)
