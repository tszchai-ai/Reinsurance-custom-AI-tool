# Reinsurance Custom AI Tool

Supporting materials for **“Custom AI Development for Reinsurance Pricing Workflow”**, a research project by Phoebe Choi and Tsz Chai Fung. The project develops two complementary Custom GPT tools:

1. **Broker submission validation** compares consecutive-year reinsurance submissions and localizes pricing-relevant inconsistencies in claims, policies, premium, rate change, exposure, treaty structure, and loss triangles.
2. **Experience rating** turns a validated submission into auditable actuarial calculations using a deterministic Python function library for premium on-leveling, trend, development, loss ratios, credibility, and claim-level excess-layer analysis.

The workbooks in this repository are synthetic research examples. Results support actuarial review and judgment; they are not autonomous pricing or quoting decisions.

## Access the Custom GPTs and Results

### Broker Submission Validation Custom GPT

- [Open the Broker Submission Validation Custom GPT](https://chatgpt.com/g/g-6a726bc3377481919161a1011336d446-reinsurance-custom-gpt)

| Test | Generic ChatGPT result | Custom GPT result |
|---|---|---|
| Energy quota share (QS) | [Without the Custom GPT](https://chatgpt.com/share/6a711141-beec-83ea-814d-8c4ad49e479c) | [With the Custom GPT](https://chatgpt.com/share/6a72c40b-a4ec-83ea-8e1b-6ea56d44a223) |
| Umbrella excess of loss (XoL) | [Without the Custom GPT](https://chatgpt.com/share/6a714792-b274-83ea-aece-cf58d4d50ada) | [With the Custom GPT](https://chatgpt.com/share/6a72c708-7c10-83ea-8e1b-6ea56d44a223) |
| Property XoL | [Without the Custom GPT](https://chatgpt.com/share/6a72a69e-1ce8-83ea-aec8-50968f09eadf) | [Standard thinking](https://chatgpt.com/share/6a736f7e-4f64-83ea-8bbd-ba10f825319b) · [Extended thinking](https://chatgpt.com/share/6a736a71-9654-83ea-9d4d-e79d88add773) |

The comparison runs above used the GPT-5.6 (Sol) model.

### Experience Rating Custom GPT

- [Open the Experience Rating Custom GPT](https://chatgpt.com/g/g-6a78ead7c85c8191ba248770070b989d-reinsurance-custom-gpt-experience-rating)
- [Energy QS result — standard thinking](https://chatgpt.com/share/6a7aab5e-f0ac-83ea-8fe3-7185fd9becbe)
- [Property XoL result — extended thinking](https://chatgpt.com/share/6a7aa1d1-ef98-83ea-b728-eb37bb454921)

## Repository Contents

### `broker submission validation/`

| File | Purpose |
|---|---|
| `Developer_Instruction.txt` | Authoritative developer instruction and actuarial rule catalog for the validation Custom GPT. It defines the prior-versus-current comparison workflow, severity and tolerance rules, claim/policy/premium/triangle checks, precise workbook provenance, and synchronized report and JSON outputs. This file is uploaded to the Custom GPT knowledge base. |
| `Instruction box.txt` | Concise instruction-box prompt used to configure the validation Custom GPT. It establishes the assistant’s role, non-negotiable check priorities, operating sequence, and required output format while deferring to `Developer_Instruction.txt` for the full specification. |
| `Example1_QS_Meta.xlsx` | Few-shot meta-workbook for quota-share submissions. It packages multiple prefixed prior/current Energy QS examples covering premium, rate change, exposures, claims, and loss triangles for use in the validation Custom GPT knowledge base. |
| `Example2_UmbrellaXoL_Meta.xlsx` | Few-shot meta-workbook for multi-subline Umbrella XoL submissions. It packages multiple prefixed prior/current examples with written premium, large claims, rate change, limit profiles, and underlying limits for the validation Custom GPT knowledge base. |
| `sim_sample1_QS.R` | R code that generates the synthetic quota-share broker-submission examples, including policies/exposures, premium, claims, rate change, and loss triangles. |
| `sim_sample2_XoL.R` | R code that generates the synthetic multi-subline XoL broker-submission examples, including written premium by state and segment, large claims, rate change, and limit profiles. |
| `test1_EnergyQS_2024.xlsx` | Prior-year synthetic Energy QS broker submission used in the out-of-sample validation test. |
| `test1_EnergyQS_2025.xlsx` | Current-year Energy QS test submission containing controlled inconsistencies for the Custom GPT evaluation. |
| `test1_EnergyQS_2025_correct.xlsx` | Clean/correct 2025 Energy QS reference workbook before the controlled test inconsistencies are applied. |
| `test1_UmbrellaXoL_2024.xlsx` | Prior-year synthetic Umbrella XoL broker submission used in the out-of-sample validation test. |
| `test1_UmbrellaXoL_2025.xlsx` | Current-year Umbrella XoL test submission containing controlled inconsistencies for the Custom GPT evaluation. |
| `test1_UmbrellaXoL_2025_correct.xlsx` | Clean/correct 2025 Umbrella XoL reference workbook before the controlled test inconsistencies are applied. |
| `test1_PropertyXoL_2024.xlsx` | Prior-year synthetic Property per-risk XoL broker submission used in the out-of-sample validation test. |
| `test1_PropertyXoL_2025.xlsx` | Current-year Property XoL test submission containing controlled inconsistencies for the Custom GPT evaluation. |
| `test1_PropertyXoL_2025_correct.xlsx` | Clean/correct 2025 Property XoL reference workbook before the controlled test inconsistencies are applied. |
| `result_test1_EnergyQS.txt` | Complete machine-readable JSON issues register produced by the validation Custom GPT for the Energy QS test. |
| `result_test1_UmbrellaXoL.txt` | Complete machine-readable JSON issues register produced by the validation Custom GPT for the Umbrella XoL test. |
| `result_test1_PropertyXoL.txt` | Complete machine-readable JSON issues register from the standard-thinking Property XoL validation run. |
| `result_test1_PropertyXoL_extend.txt` | Complete machine-readable JSON issues register from the extended-thinking Property XoL validation run. |

### `experience rating/`

| File | Purpose |
|---|---|
| `Developer_Instruction_rating.txt` | Authoritative behavioral and actuarial specification for the experience-rating Custom GPT. It governs workbook profiling, source hierarchy, method feasibility, clarification, plug-ins, unit/basis checks, validation controls, provenance, and final report/JSON structure. This file is uploaded to the Custom GPT knowledge base. |
| `Instruction box_rating.txt` | Concise instruction-box prompt used to configure the experience-rating Custom GPT, including the initial-turn contract, function-library requirement, clarification logic, and final deliverables. |
| `function_call.py` | Deterministic calculation library used by the rating Custom GPT. Its 59 functions cover rate level and on-level factors, premium and loss trend, triangle operations, Chain Ladder, Bornhuetter-Ferguson, Cape Cod and related methods, credibility, loss-ratio summaries, and claim/occurrence layer analysis. |
| `test1_EnergyQS_rating.xlsx` | Energy QS experience-rating input workbook with premium, rate change, exposure, claim, and cumulative loss-triangle data; used for Chain Ladder and Bornhuetter-Ferguson analysis brought to 2025 levels. |
| `test1_PropertyXoL_rating.xlsx` | Property per-risk XoL experience-rating input workbook with 4,136 individual claims, Commercial and Personal business, premium/rate data, and three treaty layers; used for claim-level trend, development, layering, and aggregation. |
| `Rating_Function_Calls (Energy QS).json` | Full audit register for the Energy QS rating result, including workbook profile, mapped inputs, one documented plug-in, 60 deterministic function calls, raw outputs, validation checks, results, and limitations. |
| `Rating_Function_Calls (Property XoL).json` | Full audit register for the Property XoL rating result, including workbook profile, mapped inputs, documented plug-ins, 186 deterministic function calls, raw outputs, validation checks, results, warnings, and limitations. |
| `Rating_Function_Calls_old.json` | Earlier Energy QS function-call register retained for version/reference comparison; it records 65 calls under an earlier result structure and assumption set. |
| `access_rating.txt` | Plain-text convenience list of the experience-rating Custom GPT and shared result links. The same links are presented above. |

## Suggested Workflow

1. Validate a prior/current submission pair with the Broker Submission Validation Custom GPT.
2. Review the executive report and complete JSON issues register; resolve material submission issues.
3. Analyze the cleaned current submission with the Experience Rating Custom GPT.
4. Review the computed results together with the complete function-call register, assumptions, checks, warnings, and limitations.

## Important Notes

- The workbooks and results are research examples designed to demonstrate reproducible workflows.
- Custom GPT and shared-chat links may require a compatible ChatGPT account and the permissions set by their owners.
- Final assumptions, method selections, underwriting decisions, and prices remain the responsibility of qualified actuarial and reinsurance professionals.
