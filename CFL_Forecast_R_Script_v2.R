# =============================================================================
# CISCO FORECAST LEAGUE (CFL) — PHASE 1  |  R Implementation
# Target Quarter  : FY26 Q2 (November 2025 – January 2026)
# Methodology     : Aligned with "Forecasting 101" (Cisco, March 2026)
#                   Slides 5, 8, 9, 11, 12
# =============================================================================
# FRAMEWORK (from PDF)
#   Slide 9  → "Forecasting is ART + SCIENCE"
#              Statistical models (Science, 40%) + Expert consensus (Art, 60%)
#   Slide 11 → Time-series taxonomy: Stable / Growing / Declining /
#              Seasonal / Volatile / NPI
#   Slide 12 → Method selection per type: ES, Holt, ARIMA, Regression, etc.
#   Slide 5  → Cisco baseline: 4-quarter average + Seasonal index
#   Slide 8  → Accuracy = max(0, 1 − |Bias|),  Bias = (Fcst−Act)/Act
# =============================================================================

# ── 0. PACKAGES ───────────────────────────────────────────────────────────────
pkgs <- c("readxl","dplyr","tidyr","ggplot2","scales",
          "patchwork","openxlsx","RColorBrewer","gridExtra")
invisible(lapply(pkgs, function(p) {
  if (!requireNamespace(p, quietly=TRUE)) install.packages(p)
  library(p, character.only=TRUE)
}))

# ── 1. DATA LOADING ───────────────────────────────────────────────────────────
DATA_FILE <- "CFL_External_Data_Pack_Phase1.xlsx"   # adjust path as needed

HIST_QTRS <- c("FY23Q2","FY23Q3","FY23Q4","FY24Q1","FY24Q2","FY24Q3",
               "FY24Q4","FY25Q1","FY25Q2","FY25Q3","FY25Q4","FY26Q1")

# Main bookings — skip 2 header rows
raw <- read_excel(DATA_FILE, sheet="Data Pack - Actual Bookings",
                  col_names=FALSE, skip=2)
raw[[1]] <- suppressWarnings(as.integer(raw[[1]]))
prods_raw <- raw[!is.na(raw[[1]]) & raw[[1]] %in% 1:30, ]

products <- data.frame(
  Rank    = as.integer(prods_raw[[1]]),
  Product = as.character(prods_raw[[2]]),
  PLC     = as.character(prods_raw[[3]]),
  stringsAsFactors = FALSE
)
hist_mat <- as.data.frame(lapply(4:15, function(i)
  suppressWarnings(as.numeric(as.character(prods_raw[[i]])))))
colnames(hist_mat) <- HIST_QTRS
products <- cbind(products, hist_mat)
products$Ref_DP  <- suppressWarnings(as.numeric(as.character(prods_raw[[17]])))
products$Ref_MKT <- suppressWarnings(as.numeric(as.character(prods_raw[[18]])))
products$Ref_DS  <- suppressWarnings(as.numeric(as.character(prods_raw[[19]])))

cat("✓ Loaded", nrow(products), "products\n")

# Accuracy data
acc_raw <- read_excel(DATA_FILE, sheet="Data Pack - Actual Bookings",
                      col_names=FALSE, skip=37)
acc_raw[[1]] <- suppressWarnings(as.integer(acc_raw[[1]]))
acc_raw <- acc_raw[!is.na(acc_raw[[1]]) & acc_raw[[1]] %in% 1:30, ]

safe_num <- function(x) {
  v <- suppressWarnings(as.numeric(x)); v[v < 0 | is.na(v)] <- NA; v
}
smean <- function(...) {
  vals <- c(...); vals <- vals[!is.na(vals)]; if(length(vals)==0) 0.5 else mean(vals)
}

acc <- data.frame(
  Rank    = as.integer(acc_raw[[1]]),
  DP_Q1   = safe_num(acc_raw[[3]]),  DP_Q4  = safe_num(acc_raw[[5]]),  DP_Q3  = safe_num(acc_raw[[7]]),
  MKT_Q1  = safe_num(acc_raw[[10]]), MKT_Q4 = safe_num(acc_raw[[12]]), MKT_Q3 = safe_num(acc_raw[[14]]),
  DS_Q1   = safe_num(acc_raw[[17]]), DS_Q4  = safe_num(acc_raw[[19]]), DS_Q3  = safe_num(acc_raw[[21]])
)
acc$DP_mean  <- mapply(smean, acc$DP_Q1,  acc$DP_Q4,  acc$DP_Q3)
acc$MKT_mean <- mapply(smean, acc$MKT_Q1, acc$MKT_Q4, acc$MKT_Q3)
acc$DS_mean  <- mapply(smean, acc$DS_Q1,  acc$DS_Q4,  acc$DS_Q3)

g <- c(mean(acc$DP_mean), mean(acc$MKT_mean), mean(acc$DS_mean))
g <- g / sum(g)
w_dp <- g[1]; w_mkt <- g[2]; w_ds <- g[3]
cat(sprintf("Reference weights → DP: %.3f  MKT: %.3f  DS: %.3f\n", w_dp, w_mkt, w_ds))

# ── 2. TIME-SERIES CLASSIFICATION  (Cisco Slide 11) ──────────────────────────
classify_series <- function(hist_vec, plc) {
  data <- hist_vec[!is.na(hist_vec) & hist_vec >= 0]
  n    <- length(data)
  
  if (n <= 3 || plc == "NPI-Ramp") return("NPI")
  if (plc == "Decline") return("Declining")
  
  x         <- seq_along(data)
  slope     <- coef(lm(data ~ x))[2]
  mean_val  <- mean(data)
  rel_slope <- if (mean_val > 0) slope / mean_val else 0
  cv        <- sd(data) / mean_val
  
  # Q2 indices: 1, 5, 9 (1-indexed in the 12-quarter series)
  q2_vals <- data[c(1,5,9)[c(1,5,9) <= n]]
  q3_vals <- data[c(2,6,10)[c(2,6,10) <= n]]
  seas_spread <- if (length(q2_vals)>=2 && length(q3_vals)>=2 && mean_val>0)
    abs(mean(q2_vals) - mean(q3_vals)) / mean_val else 0
  
  if      (rel_slope >  0.04) "Growing"
  else if (rel_slope < -0.04) "Declining"
  else if (seas_spread > 0.15) "Seasonal"
  else if (cv > 0.35)          "Volatile"
  else                         "Stable"
}

# ── 3. FORECASTING FUNCTIONS  (Cisco Slides 5, 12) ───────────────────────────

# --- Cisco Slide 5: 4-Quarter Average ---
four_qtr_avg <- function(data) mean(tail(data, 4))

# --- Cisco Slide 5: Seasonal Index (same-quarter trend) ---
# FY26 Q2 is the 9th Q2; historical Q2s sit at positions 1, 5, 9 in our series
seasonal_index_fcst <- function(hist_vec) {
  q2_positions <- c(1, 5, 9)
  vals <- hist_vec[q2_positions[q2_positions <= length(hist_vec)]]
  vals <- vals[!is.na(vals) & vals >= 0]
  if (length(vals) < 2) return(NA)
  x <- seq_along(vals)
  cf <- coef(lm(vals ~ x))
  max(0, cf[1] + cf[2] * (length(vals) + 1))
}

# --- Holt's Double Exponential Smoothing (Slide 12) ---
holt_linear <- function(series, alpha=0.30, beta=0.05, h=1) {
  series <- series[!is.na(series) & series >= 0]
  n <- length(series)
  if (n < 2) return(series[n])
  m <- coef(lm(series ~ seq_along(series)))[2]   # OLS initial trend
  L <- series[1]; B <- m
  for (v in series[-1]) {
    L0 <- L; B0 <- B
    L <- alpha * v + (1-alpha) * (L0+B0)
    B <- beta  * (L-L0) + (1-beta) * B0
  }
  max(0, L + h*B)
}

# --- Simple Exponential Smoothing (Slide 12: SES) ---
ses <- function(series, alpha=0.25) {
  series <- series[!is.na(series) & series >= 0]
  s <- series[1]
  for (v in series[-1]) s <- alpha*v + (1-alpha)*s
  max(0, s)
}

# --- Linear Extrapolation (Slide 5: Trend) ---
linear_extrap <- function(series, h=1) {
  series <- series[!is.na(series) & series >= 0]
  x <- seq_along(series)
  cf <- coef(lm(series ~ x))
  max(0, cf[1] + cf[2] * (length(series) + h))
}

# ── 4. METHOD DISPATCHER  (Cisco Slide 11-12 taxonomy) ───────────────────────
stat_forecast <- function(hist_vec, series_type) {
  data <- hist_vec[!is.na(hist_vec) & hist_vec >= 0]
  n    <- length(data)
  if (n == 0) return(NA)
  if (n == 1) return(data[1])
  
  fa   <- four_qtr_avg(data)
  lt   <- linear_extrap(data)
  holt <- holt_linear(data)
  si   <- seasonal_index_fcst(hist_vec)
  if (is.na(si)) si <- fa   # fallback if not enough Q2 history
  
  switch(series_type,
    # Cisco Slide 12 method assignment:
    "Stable"    = 0.40*si + 0.35*fa + 0.25*holt,          # SES / Seasonal Naive
    "Growing"   = 0.40*holt + 0.35*si + 0.25*lt,          # Holt's (captures trend acceleration)
    "Declining" = 0.55*lt + 0.45*ses(data, alpha=0.20),   # Linear Trend + dampened SES
    "Seasonal"  = 0.50*si + 0.30*holt + 0.20*fa,          # Seasonal index dominates
    "Volatile"  = 0.50*ses(data, alpha=0.35) + 0.50*fa,   # SES absorbs noise
    "NPI"       = {                                         # Short series
      if (n >= 4) 0.50*holt_linear(data, alpha=0.50, beta=0.30) + 0.50*lt
      else        holt_linear(data, alpha=0.50, beta=0.30)
    },
    fa   # default fallback
  )
}

# ── 5. COMPUTE ALL 30 FORECASTS ───────────────────────────────────────────────
results <- products %>%
  rowwise() %>%
  mutate(
    Series_Type = classify_series(c_across(all_of(HIST_QTRS)), PLC),
    Stat_Model  = stat_forecast(c_across(all_of(HIST_QTRS)), Series_Type),
    Ref_Ensemble = {
      refs <- c(Ref_DP, Ref_MKT, Ref_DS)
      wts  <- c(w_dp, w_mkt, w_ds)
      valid <- !is.na(refs)
      if (any(valid)) weighted.mean(refs[valid], wts[valid]) else Stat_Model
    },
    # Cisco Slide 9: Science (40%) + Art (60%)
    My_Forecast_FY26Q2 = round(
      ifelse(is.na(Stat_Model),
             Ref_Ensemble,
             0.40 * Stat_Model + 0.60 * Ref_Ensemble)
    )
  ) %>%
  ungroup()

# Print summary
cat("\n=== FINAL FORECAST TABLE — FY26 Q2 ===\n")
results %>%
  select(Rank, Series_Type, Product, My_Forecast_FY26Q2, Ref_DP, Ref_MKT, Ref_DS) %>%
  mutate(across(where(is.numeric), ~formatC(round(.), big.mark=","))) %>%
  print(n=30)

# ── 6. EXPORT TO EXCEL ────────────────────────────────────────────────────────
wb_out <- createWorkbook()
addWorksheet(wb_out, "My_Forecast_FY26Q2")
addWorksheet(wb_out, "Full_Data")
addWorksheet(wb_out, "Methodology")

# Submission sheet
submission <- results %>%
  select(Rank, Product, PLC, Series_Type, My_Forecast_FY26Q2, Ref_DP, Ref_MKT, Ref_DS)
writeData(wb_out, "My_Forecast_FY26Q2", submission)
addStyle(wb_out, "My_Forecast_FY26Q2",
         style = createStyle(fgFill="#1D3557", fontColour="white", textDecoration="bold",
                             halign="center", fontSize=11),
         rows=1, cols=1:8, gridExpand=TRUE)
addStyle(wb_out, "My_Forecast_FY26Q2",
         style = createStyle(fontColour="#E63946", textDecoration="bold", numFmt="#,##0"),
         rows=2:(nrow(submission)+1), cols=5, gridExpand=TRUE)
setColWidths(wb_out, "My_Forecast_FY26Q2", cols=1:8,
             widths=c(5, 46, 13, 13, 20, 15, 15, 15))

# Full data sheet
writeData(wb_out, "Full_Data",
          results %>% select(Rank, Product, PLC, Series_Type,
                             Stat_Model, Ref_Ensemble, Ref_DP, Ref_MKT, Ref_DS,
                             My_Forecast_FY26Q2))

# Methodology sheet
meth <- data.frame(
  Step = c("1","1","1","1","1","1","1","1","","2","2","2","2","","3","","Metric"),
  Concept = c(
    "Time-series classification (Slide 11)",
    "  Stable: 4-Qtr Avg + Seasonal Index + Holt's ES",
    "  Growing: Holt's ES + Seasonal Index + Linear Trend",
    "  Declining: Linear Trend + SES (dampened)",
    "  Seasonal: Seasonal Index + Holt's + 4-Qtr Avg",
    "  Volatile: SES (alpha=0.35) + 4-Qtr Avg",
    "  NPI: Holt's (aggressive) + Linear Trend",
    "Cisco baseline: 4-Quarter Average & Seasonal Index (Slide 5)",
    "",
    "Reference Ensemble (Slide 9 — Art)",
    "  Demand Planner — weight ≈ 35.7%",
    "  Marketing Team — weight ≈ 31.2%",
    "  Data Science   — weight ≈ 33.1%",
    "Weights derived from historical accuracy over last 3 quarters",
    "Final Blend: 40% Statistical + 60% Reference Ensemble (Slide 9)",
    "",
    "Accuracy = max(0, 1-|Bias|),  Bias = (Fcst-Act)/Act  (Slide 8)"
  ),
  stringsAsFactors = FALSE
)
writeData(wb_out, "Methodology", meth)
setColWidths(wb_out, "Methodology", cols=1:2, widths=c(8, 70))

saveWorkbook(wb_out, "CFL_My_Forecasts_FY26Q2.xlsx", overwrite=TRUE)
cat("\n✓ Excel saved: CFL_My_Forecasts_FY26Q2.xlsx\n")

# ── 7. VISUALISATIONS ─────────────────────────────────────────────────────────
theme_cfl <- theme_minimal(base_size=10) +
  theme(
    plot.background  = element_rect(fill="white", colour=NA),
    panel.background = element_rect(fill="#F8F9FA", colour=NA),
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(colour="#E8E8E8"),
    plot.title       = element_text(face="bold", size=11),
    plot.subtitle    = element_text(colour="#555555", size=8.5),
    legend.position  = "bottom"
  )

TYPE_COLORS <- c(Stable="#2A9D8F", Growing="#06D6A0", Declining="#E76F51",
                 Seasonal="#FFB703", Volatile="#FB8500", NPI="#8338EC")

# --- 7a. Per-product time-series (top 9) ---
ALL_QTRS <- c(HIST_QTRS, "FY26Q2")
ts_long <- results %>%
  select(Rank, Product, PLC, Series_Type, all_of(HIST_QTRS),
         My_Forecast_FY26Q2, Ref_DP, Ref_MKT, Ref_DS) %>%
  pivot_longer(all_of(HIST_QTRS), names_to="Quarter", values_to="Actuals") %>%
  mutate(Quarter = factor(Quarter, levels=HIST_QTRS))

plot_product <- function(rank_id) {
  p_row <- results %>% filter(Rank == rank_id)
  p_ts  <- ts_long  %>% filter(Rank == rank_id)
  stype <- p_row$Series_Type[1]
  
  pts <- data.frame(
    Quarter = "FY26Q2",
    Source  = c("My Forecast","Demand Planner","Marketing","Data Science"),
    Value   = c(p_row$My_Forecast_FY26Q2, p_row$Ref_DP, p_row$Ref_MKT, p_row$Ref_DS)
  )
  
  ggplot(p_ts, aes(x=Quarter, y=Actuals)) +
    geom_line(aes(group=1), colour="#1D3557", linewidth=0.9) +
    geom_point(colour="#1D3557", size=2.5, na.rm=TRUE) +
    geom_vline(xintercept=12.5, linetype="dashed", colour="#AAAAAA", linewidth=0.7) +
    geom_point(data=pts, aes(x=13, y=Value, colour=Source, shape=Source),
               size=4.5, inherit.aes=FALSE) +
    scale_colour_manual(values=c(
      "My Forecast"="#E63946","Demand Planner"="#457B9D",
      "Marketing"="#2A9D8F","Data Science"="#E9C46A")) +
    scale_shape_manual(values=c("My Forecast"=8,"Demand Planner"=17,
                                "Marketing"=15,"Data Science"=18)) +
    scale_y_continuous(labels=comma) +
    scale_x_discrete(drop=FALSE,
                     limits=c(HIST_QTRS,"FY26Q2"),
                     labels=c(HIST_QTRS,"FY26Q2\n▶Fcst")) +
    labs(title=paste0("[",rank_id,"] ",substr(p_row$Product[1],1,40)),
         subtitle=paste0(stype," | PLC: ",p_row$PLC[1]),
         x=NULL, y="Units", colour=NULL, shape=NULL) +
    theme_cfl +
    theme(axis.text.x=element_text(angle=45, hjust=1, size=7),
          plot.subtitle=element_text(
            colour=TYPE_COLORS[stype], face="italic", size=8))
}

p_grid <- wrap_plots(lapply(1:9, plot_product), ncol=3) +
  plot_annotation(
    title    = "CFL FY26 Q2 — Top 9 Products: Time-Series & Forecasts",
    subtitle = paste0("Cisco Forecasting 101 | Series-type aware methods (Slide 11-12) | ",
                      "Blend: 40% Statistical + 60% Reference Ensemble"),
    theme    = theme(plot.title=element_text(face="bold",size=14))
  )
ggsave("CFL_Top9_TimeSeries.png", p_grid, width=19, height=15, dpi=150)
cat("✓ Fig 1 saved: CFL_Top9_TimeSeries.png\n")

# --- 7b. All-product forecast comparison bar ---
bar_df <- results %>%
  select(Rank, Product, Series_Type, My_Forecast_FY26Q2, Ref_DP, Ref_MKT, Ref_DS) %>%
  pivot_longer(c(My_Forecast_FY26Q2,Ref_DP,Ref_MKT,Ref_DS),
               names_to="Source", values_to="Units") %>%
  mutate(
    Source = recode(Source,
      My_Forecast_FY26Q2="My Forecast", Ref_DP="Demand Planner",
      Ref_MKT="Marketing", Ref_DS="Data Science"),
    Label  = paste0("[",Rank,"] ",substr(Product,1,28))
  )

p_bar <- ggplot(bar_df, aes(x=reorder(Label,-Units), y=Units, fill=Source)) +
  geom_col(position="dodge", width=0.72, alpha=0.88) +
  scale_fill_manual(values=c("My Forecast"="#E63946","Demand Planner"="#457B9D",
                             "Marketing"="#2A9D8F","Data Science"="#E9C46A")) +
  scale_y_continuous(labels=comma) +
  labs(title="FY26 Q2 Forecast Comparison — All 30 Cisco Products",
       subtitle="My Forecast (★) vs DP / Marketing / Data Science reference forecasts",
       x=NULL, y="Forecasted Units (FY26 Q2)", fill=NULL) +
  theme_cfl +
  theme(axis.text.x=element_text(angle=60, hjust=1, size=7.5))
ggsave("CFL_Forecast_Comparison.png", p_bar, width=22, height=9, dpi=150)
cat("✓ Fig 2 saved: CFL_Forecast_Comparison.png\n")

# --- 7c. Accuracy heatmap ---
heat_df <- acc %>%
  left_join(results %>% select(Rank,Product), by="Rank") %>%
  pivot_longer(c(DP_mean,MKT_mean,DS_mean), names_to="Team", values_to="Accuracy") %>%
  mutate(
    Team  = recode(Team, DP_mean="Demand Planner", MKT_mean="Marketing", DS_mean="Data Science"),
    Label = paste0("[",Rank,"] ",substr(Product,1,33))
  )

p_heat <- ggplot(heat_df, aes(x=Team, y=reorder(Label,-Rank), fill=Accuracy)) +
  geom_tile(colour="white", linewidth=0.4) +
  geom_text(aes(label=sprintf("%.2f",Accuracy)), size=2.9, colour="black") +
  scale_fill_gradientn(colours=c("#E63946","#FFF3CD","#2A9D8F"),
                       limits=c(0.30,1.00), name="Accuracy") +
  labs(title="Historical Forecast Accuracy by Reference Team",
       subtitle="Avg over FY25Q3 / FY25Q4 / FY26Q1 — drives ensemble weights",
       x=NULL, y=NULL) +
  theme_cfl + theme(axis.text.y=element_text(size=7.5))
ggsave("CFL_Accuracy_Heatmap.png", p_heat, width=10, height=14, dpi=150)
cat("✓ Fig 3 saved: CFL_Accuracy_Heatmap.png\n")

# --- 7d. Series-type taxonomy + top 10 ---
p_pie_df <- results %>% count(Series_Type) %>%
  mutate(pct=n/sum(n), label=paste0(Series_Type,"\n(",n,")"))

p_pie <- ggplot(p_pie_df, aes(x="",y=n,fill=Series_Type)) +
  geom_col(width=1, colour="white") + coord_polar(theta="y") +
  geom_text(aes(label=paste0(round(pct*100),"%")),
            position=position_stack(vjust=0.5), fontface="bold", size=4) +
  scale_fill_manual(values=TYPE_COLORS) +
  labs(title="Time-Series Classification\n(Cisco Slide 11 Taxonomy)", fill=NULL) +
  theme_void() + theme(plot.title=element_text(face="bold",hjust=0.5,size=12))

top10_df <- results %>% slice_max(My_Forecast_FY26Q2,n=10) %>%
  mutate(Label=paste0("[",Rank,"] ",substr(Product,1,33)))

p_top10 <- ggplot(top10_df, aes(x=reorder(Label,My_Forecast_FY26Q2),
                                  y=My_Forecast_FY26Q2, fill=Series_Type)) +
  geom_col(alpha=0.88, width=0.7) +
  geom_text(aes(label=comma(My_Forecast_FY26Q2)), hjust=-0.08, size=3.3) +
  scale_fill_manual(values=TYPE_COLORS) +
  scale_y_continuous(labels=comma, expand=expansion(mult=c(0,0.15))) +
  coord_flip() +
  labs(title="Top 10 Products by Forecast Volume\n(Colour = Series Type)",
       x=NULL, y="My Forecast FY26Q2 (Units)", fill=NULL) +
  theme_cfl

p_combined <- p_pie + p_top10 +
  plot_annotation(title="CFL FY26 Q2 — Portfolio Overview",
                  theme=theme(plot.title=element_text(face="bold",size=14)))
ggsave("CFL_Portfolio_Overview.png", p_combined, width=17, height=7, dpi=150)
cat("✓ Fig 4 saved: CFL_Portfolio_Overview.png\n")

# ── 8. METHODOLOGY BOX ────────────────────────────────────────────────────────
cat("\n")
cat("╔═══════════════════════════════════════════════════════════════════╗\n")
cat("║     CFL FORECAST METHODOLOGY  —  Cisco Forecasting 101 v2       ║\n")
cat("╠═══════════════════════════════════════════════════════════════════╣\n")
cat("║ SOURCE        │ SLIDE │ APPROACH                                 ║\n")
cat("╠═══════════════════════════════════════════════════════════════════╣\n")
cat("║ Classification│  11   │ Stable / Growing / Declining /           ║\n")
cat("║               │       │ Seasonal / Volatile / NPI                ║\n")
cat("║ 4-Qtr Average │   5   │ avg(FY25Q1 – FY26Q1) — Cisco baseline   ║\n")
cat("║ Seasonal Index│   5   │ OLS on Q2s (FY23Q2, FY24Q2, FY25Q2)    ║\n")
cat("║ Holt's ES     │  12   │ alpha=0.30, beta=0.05; OLS-init trend   ║\n")
cat("║ Linear Trend  │   5   │ OLS extrap — Declining & NPI products   ║\n")
cat("╠═══════════════════════════════════════════════════════════════════╣\n")
cat("║ Reference     │   9   │ Art — weighted by historical accuracy    ║\n")
cat(sprintf("║ Ensemble      │       │ DP=%.3f | MKT=%.3f | DS=%.3f         ║\n", w_dp, w_mkt, w_ds))
cat("╠═══════════════════════════════════════════════════════════════════╣\n")
cat("║ Final Blend   │   9   │ 40% Science + 60% Art                   ║\n")
cat("║ Accuracy KPI  │   8   │ max(0, 1 − |Bias|)                      ║\n")
cat("╚═══════════════════════════════════════════════════════════════════╝\n")

cat("\n✅ Deliverables:\n")
cat("  CFL_My_Forecasts_FY26Q2.xlsx   — submission-ready (3 sheets)\n")
cat("  CFL_Top9_TimeSeries.png        — trend + forecast plot, top 9\n")
cat("  CFL_Forecast_Comparison.png    — all 30 products, 4-way bar chart\n")
cat("  CFL_Accuracy_Heatmap.png       — team accuracy heatmap\n")
cat("  CFL_Portfolio_Overview.png     — taxonomy pie + top 10 bar\n")
