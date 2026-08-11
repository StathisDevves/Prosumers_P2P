# Blockchain-Enabled Local Energy Markets for Climate Resilience

This repository contains the modelling framework, data processing workflow, and comparative results developed for the study:

**“Blockchain-enabled Local Energy Markets as Climate Resilience Infrastructure: A Welfare-Based Framework for Distributed Flexibility, Storage and Community Energy Trading.”**

The project develops an integrated 8,760-hour residential prosumer model combining photovoltaic generation, battery energy storage, dynamic electricity pricing, demand response, endogenous utility functions, grid interaction, peer-to-peer (P2P) energy trading, and welfare analysis.

Three battery management strategies are compared under four storage capacities (10, 15, 20, and 30 kWh):

- **Strategy A:** fixed rule-based battery dispatch;
- **Strategy B:** adaptive forecast-based threshold control;
- **Strategy C:** 24-hour rolling-horizon optimisation.

All strategies are evaluated under the same demand profile, PV production, electricity prices, battery constraints, utility functions, and welfare equations, enabling a consistent comparison of the effect of decision intelligence and storage capacity.

The framework evaluates energy, economic, behavioural, and market-performance indicators, including:

- grid electricity imports and exports;
- PV self-consumption;
- battery operation and storage utilisation;
- dynamic price elasticity;
- prosumer utility;
- consumer and prosumer surplus;
- conventional grid-related surplus;
- P2P revenues and P2P social welfare;
- dynamic P2P prices;
- utility-based P2P clearing prices;
- total social welfare.

A key methodological contribution is the proposed **multi-layer welfare decomposition**, which separates:

1. **Behavioural welfare** through the prosumer utility function;
2. **Conventional market welfare** through grid/spot-market prosumer surplus;
3. **Decentralised market welfare** through P2P surplus and social welfare.

The study also applies **24-hour Fourier decomposition** to elasticity, P2P price differences, and participant utility functions in order to identify systematic intraday market patterns that are not visible through annual indicators alone.

From a climate-resilience perspective, distributed storage, intelligent dispatch, and blockchain-enabled local energy trading are interpreted as resilience infrastructure capable of reducing exposure to wholesale-market volatility, increasing renewable self-consumption, improving local energy autonomy, and strengthening the adaptive capacity of energy communities.

## Repository contents

The repository may include:

- input datasets and calibrated hourly profiles;
- Python notebooks and scripts;
- Excel-based model outputs;
- comparative annual and monthly KPI analyses;
- Fourier-analysis workbooks;
- P2P clearing-price analysis;
- figures and charts used in the manuscript;
- methodological notes and supplementary material.

## Main research themes

`Prosumer Optimisation` · `Battery Energy Storage` · `Peer-to-Peer Energy Trading` · `Blockchain Energy Markets` · `Demand Response` · `Social Welfare` · `Dynamic Pricing` · `Climate Resilience`
