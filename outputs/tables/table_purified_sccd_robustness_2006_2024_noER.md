# Purified SCCD robustness check

Status: not completed from the available data and scripts.

Reason: the Stata data file contains the composite SCCD variable but does not contain the original entropy-weighted indicator-level variables needed to reconstruct SCCD after removing ER, patent-related, trade-related, fiscal-expenditure-related, or digital-overlap indicators. The current do-files also do not include the entropy-weighting construction code.

Additional check: MCCD is identical to SCCD in the available data and therefore cannot be used as a purified SCCD proxy.

Required inputs to complete this check: raw indicator-level panel, the entropy-weighting script, the mapping from each indicator to SCCD dimensions and criterion layers, and explicit rules for excluding overlap indicators.
