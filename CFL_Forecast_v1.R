# =============================================================================
# Cisco Forecast League - Forecasting Pipeline (R)
# Efficient technique: rolling-origin CV weighted ensemble of ETS/ARIMA/SES/Baseline
# =============================================================================

# ---------------------------
# 0) Package management
# ---------------------------
required_pkgs <- c("readxl", "dplyr", "purrr", "openxlsx", "tibble")

for (pkg in required_pkgs) {
	if (!requireNamespace(pkg, quietly = TRUE)) {
		install.packages(pkg, repos = "https://cloud.r-project.org")
	}
	library(pkg, character.only = TRUE)
}

# ---------------------------
# 1) Configuration
# ---------------------------
DATA_FILE <- "CFL_External Data Pack_Phase1.xlsx"
TARGET_QUARTER <- "FY26Q2"
OUTPUT_FILE <- "CFL_Forecast_Output_FY26Q2.xlsx"
BACKTEST_QTRS <- c("FY25Q3", "FY25Q4", "FY26Q1")
TUNING_QTRS <- c("FY24Q3", "FY24Q4", "FY25Q1", "FY25Q2", "FY25Q3", "FY25Q4", "FY26Q1")

# Guardrail to reduce overfitting in per-product blend tuning.
W_STAT_MIN <- 0.10
W_STAT_MAX <- 0.50

HIST_QTRS <- c(
	"FY23Q2", "FY23Q3", "FY23Q4", "FY24Q1", "FY24Q2", "FY24Q3",
	"FY24Q4", "FY25Q1", "FY25Q2", "FY25Q3", "FY25Q4", "FY26Q1"
)

# Blend weights (keep aligned with Cisco's Art + Science approach)
W_STAT <- 0.40
W_REF <- 0.60

# ---------------------------
# 2) Utilities
# ---------------------------
to_num <- function(x) {
	suppressWarnings(as.numeric(as.character(x)))
}

safe_accuracy <- function(x) {
	v <- to_num(x)
	if (is.na(v)) return(NA_real_)
	if (v > 1 && v <= 100) v <- v / 100
	v <- max(0, min(1, v))
	v
}

safe_weighted_mean <- function(values, weights, fallback = NA_real_) {
	valid <- !is.na(values) & !is.na(weights) & weights > 0
	if (!any(valid)) return(fallback)
	weighted.mean(values[valid], weights[valid])
}

mae <- function(actual, pred) {
	mean(abs(actual - pred), na.rm = TRUE)
}

calc_accuracy_metric <- function(actual, forecast) {
	if (is.na(actual) || is.na(forecast)) return(NA_real_)
	if (actual == 0) return(ifelse(abs(forecast) < 1e-9, 1, 0))
	bias <- (forecast - actual) / actual
	max(0, 1 - abs(bias))
}

sanitize_series <- function(x) {
	x[!is.na(x) & x >= 0]
}

clamp <- function(x, lo, hi) {
	pmax(lo, pmin(hi, x))
}

# ---------------------------
# 3) Read and shape data
# ---------------------------
read_cfl_data <- function(file_path) {
	raw <- read_excel(file_path, sheet = "Data Pack - Actual Bookings", col_names = FALSE, skip = 2)
	raw[[1]] <- suppressWarnings(as.integer(raw[[1]]))

	products_raw <- raw[!is.na(raw[[1]]) & raw[[1]] %in% 1:30, ]

	hist_mat <- as.data.frame(lapply(4:15, function(i) to_num(products_raw[[i]])))
	colnames(hist_mat) <- HIST_QTRS

	products <- tibble(
		Rank = as.integer(products_raw[[1]]),
		Product = as.character(products_raw[[2]]),
		PLC = as.character(products_raw[[3]])
	) %>%
		bind_cols(hist_mat) %>%
		mutate(
			Ref_DP = to_num(products_raw[[17]]),
			Ref_MKT = to_num(products_raw[[18]]),
			Ref_DS = to_num(products_raw[[19]])
		) %>%
		arrange(Rank) %>%
		group_by(Rank) %>%
		slice(1) %>%
		ungroup()

	# Historical team-accuracy section
	acc_raw <- read_excel(file_path, sheet = "Data Pack - Actual Bookings", col_names = FALSE, skip = 37)
	acc_raw[[1]] <- suppressWarnings(as.integer(acc_raw[[1]]))
	acc_raw <- acc_raw[!is.na(acc_raw[[1]]) & acc_raw[[1]] %in% 1:30, ]

	acc <- tibble(
		Rank = as.integer(acc_raw[[1]]),
		DP_Q1 = vapply(acc_raw[[3]], safe_accuracy, numeric(1)),
		DP_Q4 = vapply(acc_raw[[5]], safe_accuracy, numeric(1)),
		DP_Q3 = vapply(acc_raw[[7]], safe_accuracy, numeric(1)),
		MKT_Q1 = vapply(acc_raw[[10]], safe_accuracy, numeric(1)),
		MKT_Q4 = vapply(acc_raw[[12]], safe_accuracy, numeric(1)),
		MKT_Q3 = vapply(acc_raw[[14]], safe_accuracy, numeric(1)),
		DS_Q1 = vapply(acc_raw[[17]], safe_accuracy, numeric(1)),
		DS_Q4 = vapply(acc_raw[[19]], safe_accuracy, numeric(1)),
		DS_Q3 = vapply(acc_raw[[21]], safe_accuracy, numeric(1))
	) %>%
		mutate(
			DP_mean = rowMeans(across(c(DP_Q1, DP_Q4, DP_Q3)), na.rm = TRUE),
			MKT_mean = rowMeans(across(c(MKT_Q1, MKT_Q4, MKT_Q3)), na.rm = TRUE),
			DS_mean = rowMeans(across(c(DS_Q1, DS_Q4, DS_Q3)), na.rm = TRUE)
		) %>%
		arrange(Rank) %>%
		group_by(Rank) %>%
		slice(1) %>%
		ungroup()

	list(products = products, accuracy = acc)
}

# ---------------------------
# 4) Efficient statistical forecasting engine
# ---------------------------
last4_avg_forecast <- function(y) {
	if (length(y) == 0) return(NA_real_)
	mean(tail(y, min(4, length(y))))
}

ses_forecast <- function(y) {
	if (length(y) < 2) return(last4_avg_forecast(y))
	fit <- stats::HoltWinters(ts(y, frequency = 4), beta = FALSE, gamma = FALSE)
	as.numeric(stats::predict(fit, n.ahead = 1)[1])
}

ets_forecast <- function(y) {
	if (length(y) < 4) return(last4_avg_forecast(y))
	fit <- stats::HoltWinters(ts(y, frequency = 4), beta = TRUE, gamma = FALSE)
	as.numeric(stats::predict(fit, n.ahead = 1)[1])
}

arima_forecast <- function(y) {
	if (length(y) < 6) return(last4_avg_forecast(y))
	fit <- stats::arima(ts(y, frequency = 4), order = c(1, 1, 0))
	as.numeric(stats::predict(fit, n.ahead = 1)$pred[1])
}

drift_forecast <- function(y) {
	n <- length(y)
	if (n < 2) return(last4_avg_forecast(y))
	y[n] + (y[n] - y[1]) / (n - 1)
}

seasonal_naive_forecast <- function(y) {
	n <- length(y)
	if (n < 4) return(last4_avg_forecast(y))
	y[n - 3]
}

build_method_lookup <- function() {
	list(
		last4 = last4_avg_forecast,
		ses = ses_forecast,
		ets = ets_forecast,
		arima = arima_forecast,
		drift = drift_forecast,
		snaive = seasonal_naive_forecast
	)
}

safe_model_forecast <- function(fun, y) {
	out <- tryCatch(fun(y), error = function(e) NA_real_)
	if (is.na(out)) out <- last4_avg_forecast(y)
	max(0, out)
}

rolling_cv_weights <- function(y, methods, min_train = 6) {
	n <- length(y)
	method_names <- names(methods)

	if (n <= min_train) {
		w <- rep(1 / length(methods), length(methods))
		names(w) <- method_names
		return(w)
	}

	cv_errors <- lapply(methods, function(.) numeric(0))

	for (t in min_train:(n - 1)) {
		train <- y[1:t]
		actual <- y[t + 1]
		for (m in method_names) {
			pred <- safe_model_forecast(methods[[m]], train)
			cv_errors[[m]] <- c(cv_errors[[m]], abs(actual - pred))
		}
	}

	err_vec <- vapply(cv_errors, function(e) if (length(e) == 0) Inf else mean(e, na.rm = TRUE), numeric(1))

	# Convert error to weight using inverse error scoring.
	scores <- 1 / (err_vec + 1e-6)
	if (!all(is.finite(scores)) || sum(scores) == 0) {
		w <- rep(1 / length(methods), length(methods))
	} else {
		w <- scores / sum(scores)
	}
	names(w) <- method_names
	w
}

forecast_one_series <- function(hist_vec, methods, min_train = 6) {
	y <- sanitize_series(hist_vec)
	if (length(y) == 0) return(list(stat_forecast = NA_real_, weights = NA))
	if (length(y) == 1) {
		w1 <- c(1)
		names(w1) <- names(methods)[1]
		pred1 <- c(y[1])
		names(pred1) <- names(methods)[1]
		return(list(stat_forecast = y[1], weights = w1, component_forecasts = pred1))
	}

	min_train <- min(min_train, max(3, length(y) - 1))
	w <- rolling_cv_weights(y, methods, min_train = min_train)
	preds <- vapply(methods, function(fn) safe_model_forecast(fn, y), numeric(1))
	stat_fcst <- sum(w * preds)

	list(
		stat_forecast = max(0, stat_fcst),
		weights = w,
		component_forecasts = preds
	)
}

build_candidate_model_sets <- function() {
	list(
		core4 = c("last4", "ses", "ets", "arima"),
		smooth3 = c("last4", "ses", "ets"),
		trend4 = c("last4", "ets", "arima", "drift"),
		seasonal4 = c("last4", "ses", "ets", "snaive"),
		robust5 = c("last4", "ses", "ets", "arima", "drift")
	)
}

backtest_model_set <- function(products, hist_qtrs, target_indices, methods, set_name) {
	out <- vector("list", length = nrow(products) * length(target_indices))
	k <- 1

	for (i in seq_len(nrow(products))) {
		hist_all <- as.numeric(products[i, hist_qtrs])
		for (idx in target_indices) {
			actual <- hist_all[idx]
			if (is.na(actual)) next

			train <- hist_all[seq_len(idx - 1)]
			fc <- forecast_one_series(train, methods = methods, min_train = 6)$stat_forecast
			acc <- calc_accuracy_metric(actual, fc)

			out[[k]] <- tibble(
				Rank = products$Rank[i],
				Quarter = hist_qtrs[idx],
				Model_Set = set_name,
				Actual = actual,
				Stat_Pred = fc,
				Abs_Error = abs(actual - fc),
				Stat_Accuracy = acc
			)
			k <- k + 1
		}
	}

	out <- out[!vapply(out, is.null, logical(1))]
	if (length(out) == 0) return(tibble())
	bind_rows(out)
}

evaluate_candidate_sets <- function(products, hist_qtrs, target_qtrs) {
	method_lookup <- build_method_lookup()
	candidate_sets <- build_candidate_model_sets()
	target_indices <- match(target_qtrs, hist_qtrs)
	target_indices <- target_indices[!is.na(target_indices)]

	all_details <- list()
	all_summaries <- list()

	for (set_name in names(candidate_sets)) {
		method_names <- candidate_sets[[set_name]]
		methods <- method_lookup[method_names]
		detail <- backtest_model_set(products, hist_qtrs, target_indices, methods, set_name)

		if (nrow(detail) == 0) {
			summary_row <- tibble(
				Model_Set = set_name,
				Methods = paste(method_names, collapse = ", "),
				N_Points = 0,
				Mean_Abs_Error = NA_real_,
				Mean_Accuracy = NA_real_
			)
		} else {
			summary_row <- detail %>%
				summarise(
					Model_Set = set_name,
					Methods = paste(method_names, collapse = ", "),
					N_Points = dplyr::n(),
					Mean_Abs_Error = mean(Abs_Error, na.rm = TRUE),
					Mean_Accuracy = mean(Stat_Accuracy, na.rm = TRUE)
				)
		}

		all_details[[set_name]] <- detail
		all_summaries[[set_name]] <- summary_row
	}

	summary_tbl <- bind_rows(all_summaries) %>%
		arrange(desc(Mean_Accuracy), Mean_Abs_Error)

	best_set <- summary_tbl$Model_Set[1]
	selected_methods <- candidate_sets[[best_set]]
	selected_detail <- all_details[[best_set]]

	list(
		summary = summary_tbl,
		selected_set = best_set,
		selected_methods = selected_methods,
		selected_detail = selected_detail,
		details_by_set = all_details,
		candidate_sets = candidate_sets,
		method_lookup = method_lookup
	)
}

select_best_set_per_product <- function(details_by_set) {
	all_detail <- bind_rows(details_by_set)

	if (nrow(all_detail) == 0) {
		return(list(
			best_by_rank = tibble(),
			score_by_rank_set = tibble(),
			selected_detail = tibble()
		))
	}

	score_by_rank_set <- all_detail %>%
		group_by(Rank, Model_Set) %>%
		summarise(
			Mean_Accuracy = mean(Stat_Accuracy, na.rm = TRUE),
			Mean_Abs_Error = mean(Abs_Error, na.rm = TRUE),
			N_Points = dplyr::n(),
			.groups = "drop"
		)

	best_by_rank <- score_by_rank_set %>%
		arrange(Rank, desc(Mean_Accuracy), Mean_Abs_Error) %>%
		group_by(Rank) %>%
		slice(1) %>%
		ungroup()

	selected_detail <- all_detail %>%
		inner_join(best_by_rank %>% select(Rank, Model_Set), by = c("Rank", "Model_Set"))

	list(
		best_by_rank = best_by_rank,
		score_by_rank_set = score_by_rank_set,
		selected_detail = selected_detail
	)
}

build_reference_backtest_accuracy <- function(acc, ref_weights) {
	q3 <- acc %>%
		transmute(
			Rank,
			Quarter = "FY25Q3",
			DP_Accuracy = DP_Q3,
			MKT_Accuracy = MKT_Q3,
			DS_Accuracy = DS_Q3
		)

	q4 <- acc %>%
		transmute(
			Rank,
			Quarter = "FY25Q4",
			DP_Accuracy = DP_Q4,
			MKT_Accuracy = MKT_Q4,
			DS_Accuracy = DS_Q4
		)

	q1 <- acc %>%
		transmute(
			Rank,
			Quarter = "FY26Q1",
			DP_Accuracy = DP_Q1,
			MKT_Accuracy = MKT_Q1,
			DS_Accuracy = DS_Q1
		)

	bind_rows(q3, q4, q1) %>%
		mutate(
			Ref_Ensemble_Accuracy = pmap_dbl(
				list(DP_Accuracy, MKT_Accuracy, DS_Accuracy),
				~ safe_weighted_mean(
					values = c(..1, ..2, ..3),
					weights = c(ref_weights["DP"], ref_weights["MKT"], ref_weights["DS"]),
					fallback = NA_real_
				)
			)
		)
}

tune_blend_weights <- function(stat_backtest, ref_backtest, step = 0.01) {
	joined <- stat_backtest %>%
		select(Rank, Quarter, Stat_Accuracy) %>%
		left_join(ref_backtest %>% select(Rank, Quarter, Ref_Ensemble_Accuracy), by = c("Rank", "Quarter"))

	grid <- seq(0, 1, by = step)
	grid_scores <- lapply(grid, function(w_stat) {
		w_ref <- 1 - w_stat

		proxy_acc <- mapply(function(a_stat, a_ref) {
			if (is.na(a_stat) && is.na(a_ref)) return(NA_real_)
			if (is.na(a_ref)) return(a_stat)
			if (is.na(a_stat)) return(a_ref)

			e_stat <- 1 - a_stat
			e_ref <- 1 - a_ref
			e_blend <- sqrt((w_stat * e_stat)^2 + (w_ref * e_ref)^2)
			max(0, 1 - e_blend)
		}, joined$Stat_Accuracy, joined$Ref_Ensemble_Accuracy)

		tibble(
			W_STAT = w_stat,
			W_REF = w_ref,
			Proxy_Mean_Accuracy = mean(proxy_acc, na.rm = TRUE)
		)
	})

	grid_scores <- bind_rows(grid_scores) %>% arrange(desc(Proxy_Mean_Accuracy), W_STAT)
	best <- grid_scores[1, ]

	list(
		w_stat = best$W_STAT,
		w_ref = best$W_REF,
		grid_scores = grid_scores,
		joined_backtest = joined
	)
}

tune_blend_weights_by_product <- function(stat_backtest, ref_backtest, global_w_stat, step = 0.01, w_stat_min = 0.10, w_stat_max = 0.50) {
	joined <- stat_backtest %>%
		select(Rank, Quarter, Stat_Accuracy) %>%
		left_join(ref_backtest %>% select(Rank, Quarter, Ref_Ensemble_Accuracy), by = c("Rank", "Quarter"))

	w_stat_min <- max(0, min(1, w_stat_min))
	w_stat_max <- max(0, min(1, w_stat_max))
	if (w_stat_min > w_stat_max) {
		tmp <- w_stat_min
		w_stat_min <- w_stat_max
		w_stat_max <- tmp
	}

	grid <- seq(w_stat_min, w_stat_max, by = step)
	if (tail(grid, 1) < w_stat_max) grid <- c(grid, w_stat_max)
	if (length(grid) == 0) grid <- c(clamp(global_w_stat, w_stat_min, w_stat_max))
	rank_splits <- split(joined, joined$Rank)

	weights_by_rank <- bind_rows(lapply(names(rank_splits), function(rank_id) {
		df <- rank_splits[[rank_id]]

		scores <- sapply(grid, function(w_stat) {
			w_ref <- 1 - w_stat
			proxy_acc <- mapply(function(a_stat, a_ref) {
				if (is.na(a_stat) && is.na(a_ref)) return(NA_real_)
				if (is.na(a_ref)) return(a_stat)
				if (is.na(a_stat)) return(a_ref)

				e_stat <- 1 - a_stat
				e_ref <- 1 - a_ref
				e_blend <- sqrt((w_stat * e_stat)^2 + (w_ref * e_ref)^2)
				max(0, 1 - e_blend)
			}, df$Stat_Accuracy, df$Ref_Ensemble_Accuracy)

			mean(proxy_acc, na.rm = TRUE)
		})

		if (all(!is.finite(scores))) {
			best_w_stat <- global_w_stat
			best_score <- NA_real_
		} else {
			best_idx <- which.max(scores)
			best_w_stat <- grid[best_idx]
			best_score <- scores[best_idx]
		}

		tibble(
			Rank = as.integer(rank_id),
			W_STAT_Product = best_w_stat,
			W_REF_Product = 1 - best_w_stat,
			Product_Proxy_Mean_Accuracy = best_score
		)
	}))

	joined_tuned <- joined %>%
		left_join(weights_by_rank, by = "Rank") %>%
		mutate(
			W_STAT_Product = ifelse(is.na(W_STAT_Product), clamp(global_w_stat, w_stat_min, w_stat_max), W_STAT_Product),
			W_STAT_Product = clamp(W_STAT_Product, w_stat_min, w_stat_max),
			W_REF_Product = 1 - W_STAT_Product,
			Blended_Proxy_Accuracy = mapply(function(a_stat, a_ref, w_stat, w_ref) {
				if (is.na(a_stat) && is.na(a_ref)) return(NA_real_)
				if (is.na(a_ref)) return(a_stat)
				if (is.na(a_stat)) return(a_ref)

				e_stat <- 1 - a_stat
				e_ref <- 1 - a_ref
				e_blend <- sqrt((w_stat * e_stat)^2 + (w_ref * e_ref)^2)
				max(0, 1 - e_blend)
			}, Stat_Accuracy, Ref_Ensemble_Accuracy, W_STAT_Product, W_REF_Product)
		)

	overall <- joined_tuned %>%
		summarise(
			N_Backtest_Points = dplyr::n(),
			Stat_Mean_Accuracy = mean(Stat_Accuracy, na.rm = TRUE),
			Ref_Mean_Accuracy = mean(Ref_Ensemble_Accuracy, na.rm = TRUE),
			Blended_Proxy_Mean_Accuracy = mean(Blended_Proxy_Accuracy, na.rm = TRUE)
		)

	list(
		weights_by_rank = weights_by_rank,
		joined_backtest = joined_tuned,
		overall = overall
	)
}

compute_product_confidence <- function(selected_detail, tuning_qtrs) {
	if (nrow(selected_detail) == 0) return(tibble())

	n_target <- length(tuning_qtrs)

	selected_detail %>%
		group_by(Rank) %>%
		summarise(
			N_Points = dplyr::n(),
			Mean_Accuracy = mean(Stat_Accuracy, na.rm = TRUE),
			SD_Accuracy = sd(Stat_Accuracy, na.rm = TRUE),
			Mean_APE = mean(Abs_Error / pmax(Actual, 1), na.rm = TRUE),
			.groups = "drop"
		) %>%
		mutate(
			Coverage_Factor = clamp(N_Points / n_target, 0, 1),
			Stability_Factor = clamp(1 - (coalesce(SD_Accuracy, 0.25) / 0.20), 0, 1),
			Accuracy_Factor = clamp(coalesce(Mean_Accuracy, 0), 0, 1),
			Error_Factor = clamp(1 - coalesce(Mean_APE, 1), 0, 1),
			Confidence_Score = round(100 * (0.45 * Accuracy_Factor + 0.30 * Stability_Factor + 0.15 * Coverage_Factor + 0.10 * Error_Factor), 1)
		) %>%
		select(Rank, N_Points, Mean_Accuracy, SD_Accuracy, Mean_APE, Coverage_Factor, Stability_Factor, Confidence_Score)
}

# ---------------------------
# 5) Build full forecasting pipeline
# ---------------------------
run_pipeline <- function(data_file = DATA_FILE, output_file = OUTPUT_FILE) {
	if (!file.exists(data_file)) {
		stop(sprintf("Data file not found: %s", data_file))
	}

	loaded <- read_cfl_data(data_file)
	products <- loaded$products
	acc <- loaded$accuracy

	# Global team weights from historical accuracy
	global_acc <- c(
		DP = mean(acc$DP_mean, na.rm = TRUE),
		MKT = mean(acc$MKT_mean, na.rm = TRUE),
		DS = mean(acc$DS_mean, na.rm = TRUE)
	)
	global_acc[!is.finite(global_acc)] <- 0
	if (sum(global_acc) == 0) global_acc[] <- 1
	ref_weights <- global_acc / sum(global_acc)

	# 5.1) Tune statistical model set on a broader holdout window for robustness.
	set_eval <- evaluate_candidate_sets(products, HIST_QTRS, TUNING_QTRS)
	product_set_tune <- select_best_set_per_product(set_eval$details_by_set)
	selected_method_names <- set_eval$selected_methods
	selected_methods <- set_eval$method_lookup[selected_method_names]

	# 5.2) Tune final blend using available historical reference accuracies
	ref_backtest <- build_reference_backtest_accuracy(acc, ref_weights)
	ref_backtest_tuning <- ref_backtest %>% filter(Quarter %in% TUNING_QTRS)
	blend_tune <- tune_blend_weights(product_set_tune$selected_detail, ref_backtest_tuning, step = 0.01)
	w_stat_tuned <- clamp(blend_tune$w_stat, W_STAT_MIN, W_STAT_MAX)
	w_ref_tuned <- 1 - w_stat_tuned
	blend_tune_product <- tune_blend_weights_by_product(
		product_set_tune$selected_detail,
		ref_backtest_tuning,
		global_w_stat = w_stat_tuned,
		step = 0.01,
		w_stat_min = W_STAT_MIN,
		w_stat_max = W_STAT_MAX
	)
	product_confidence <- compute_product_confidence(product_set_tune$selected_detail, TUNING_QTRS)

	# Backtest validation summary for selected statistical set and tuned blend proxy.
	validation_detail <- blend_tune_product$joined_backtest %>%
		mutate(
			Blended_Proxy_Accuracy = mapply(function(a_stat, a_ref) {
				if (is.na(a_stat) && is.na(a_ref)) return(NA_real_)
				if (is.na(a_ref)) return(a_stat)
				if (is.na(a_stat)) return(a_ref)

				e_stat <- 1 - a_stat
				e_ref <- 1 - a_ref
				e_blend <- sqrt((w_stat_tuned * e_stat)^2 + (w_ref_tuned * e_ref)^2)
				max(0, 1 - e_blend)
			}, Stat_Accuracy, Ref_Ensemble_Accuracy)
		)

	validation_overall <- blend_tune_product$overall

	validation_by_quarter <- validation_detail %>%
		group_by(Quarter) %>%
		summarise(
			N_Backtest_Points = dplyr::n(),
			Stat_Mean_Accuracy = mean(Stat_Accuracy, na.rm = TRUE),
			Ref_Mean_Accuracy = mean(Ref_Ensemble_Accuracy, na.rm = TRUE),
			Blended_Proxy_Mean_Accuracy = mean(Blended_Proxy_Accuracy, na.rm = TRUE),
			.groups = "drop"
		) %>%
		arrange(Quarter)

	validation_last3 <- validation_by_quarter %>%
		filter(Quarter %in% BACKTEST_QTRS)

	cat("Reference ensemble weights:\n")
	cat(sprintf("  DP  = %.4f\n", ref_weights["DP"]))
	cat(sprintf("  MKT = %.4f\n", ref_weights["MKT"]))
	cat(sprintf("  DS  = %.4f\n\n", ref_weights["DS"]))
	cat(sprintf("Selected statistical model set: %s (%s)\n", set_eval$selected_set, paste(selected_method_names, collapse = ", ")))
	cat(sprintf("Tuned blend weights from backtest proxy: W_STAT = %.2f, W_REF = %.2f\n\n", w_stat_tuned, w_ref_tuned))
	cat(sprintf("Per-product model-set tuning active for %d products\n", nrow(product_set_tune$best_by_rank)))
	cat(sprintf("Per-product blend-weight tuning active for %d products\n\n", nrow(blend_tune_product$weights_by_rank)))
	cat(sprintf("W_STAT guardrail active: [%.2f, %.2f]\n", W_STAT_MIN, W_STAT_MAX))
	cat(sprintf("Robust holdout tuning window: %s\n\n", paste(TUNING_QTRS, collapse = ", ")))
	cat("Backtest validation (robust holdout window):\n")
	cat(sprintf("  Stat mean accuracy          : %.4f\n", validation_overall$Stat_Mean_Accuracy[1]))
	cat(sprintf("  Reference mean accuracy     : %.4f\n", validation_overall$Ref_Mean_Accuracy[1]))
	cat(sprintf("  Blended proxy mean accuracy : %.4f\n\n", validation_overall$Blended_Proxy_Mean_Accuracy[1]))
	cat("Quarter-wise blended proxy accuracy (last 3 official backtest quarters):\n")
	print(validation_last3 %>% select(Quarter, Blended_Proxy_Mean_Accuracy), n = nrow(validation_last3))
	cat("\nTop confidence products:\n")
	print(product_confidence %>%
		left_join(products %>% select(Rank, Product), by = "Rank") %>%
		arrange(desc(Confidence_Score), desc(Mean_Accuracy)) %>%
		select(Rank, Product, Confidence_Score, Mean_Accuracy, SD_Accuracy, N_Points) %>%
		slice_head(n = 10), n = 10)
	cat("\n")

	hist_matrix <- as.matrix(products[, HIST_QTRS])
	set_by_rank <- setNames(product_set_tune$best_by_rank$Model_Set, product_set_tune$best_by_rank$Rank)

	model_outputs <- vector("list", nrow(products))
	model_set_selected_product <- character(nrow(products))
	for (i in seq_len(nrow(products))) {
		rank_i <- as.character(products$Rank[i])
		set_i <- set_by_rank[[rank_i]]
		if (is.null(set_i) || is.na(set_i)) set_i <- set_eval$selected_set
		model_set_selected_product[i] <- set_i
		method_names_i <- set_eval$candidate_sets[[set_i]]
		methods_i <- set_eval$method_lookup[method_names_i]
		model_outputs[[i]] <- forecast_one_series(as.numeric(hist_matrix[i, ]), methods = methods_i, min_train = 6)
	}

	stat_fcst <- vapply(model_outputs, function(x) x$stat_forecast, numeric(1))

	extract_component <- function(mname) {
		vapply(model_outputs, function(x) {
			if (is.null(x$component_forecasts)) return(NA_real_)
			val <- x$component_forecasts[mname]
			if (length(val) == 0 || is.na(val)) return(NA_real_)
			as.numeric(val)
		}, numeric(1))
	}

	extract_weight <- function(mname) {
		vapply(model_outputs, function(x) {
			if (is.null(x$weights)) return(NA_real_)
			val <- x$weights[mname]
			if (length(val) == 0 || is.na(val)) return(NA_real_)
			as.numeric(val)
		}, numeric(1))
	}

	all_method_names <- names(set_eval$method_lookup)
	component_df <- as_tibble(setNames(lapply(all_method_names, extract_component), paste0("Stat_", toupper(all_method_names))))
	weights_df <- as_tibble(setNames(lapply(all_method_names, extract_weight), paste0("CVW_", toupper(all_method_names))))
	w_stat_map <- setNames(blend_tune_product$weights_by_rank$W_STAT_Product, blend_tune_product$weights_by_rank$Rank)
	w_stat_vec <- vapply(products$Rank, function(r) {
		val <- w_stat_map[[as.character(r)]]
		if (is.null(val) || is.na(val)) w_stat_tuned else as.numeric(val)
	}, numeric(1))
	w_stat_vec <- clamp(w_stat_vec, W_STAT_MIN, W_STAT_MAX)
	w_ref_vec <- 1 - w_stat_vec

	products_out <- bind_cols(products, component_df, weights_df) %>%
		left_join(product_confidence %>% select(Rank, Confidence_Score, Mean_Accuracy, SD_Accuracy, N_Points), by = "Rank") %>%
		mutate(
			Stat_Model = stat_fcst,
			Model_Set_Selected = model_set_selected_product,
			W_STAT_Product = w_stat_vec,
			W_REF_Product = w_ref_vec,
			Ref_Ensemble = pmap_dbl(
				list(Ref_DP, Ref_MKT, Ref_DS),
				~ safe_weighted_mean(
					values = c(..1, ..2, ..3),
					weights = c(ref_weights["DP"], ref_weights["MKT"], ref_weights["DS"]),
					fallback = NA_real_
				)
			),
			Final_Forecast = round(
				ifelse(
					is.na(Ref_Ensemble) & !is.na(Stat_Model),
					Stat_Model,
					ifelse(is.na(Stat_Model) & !is.na(Ref_Ensemble), Ref_Ensemble, W_STAT_Product * Stat_Model + W_REF_Product * Ref_Ensemble)
				)
			)
		)

	# Output workbook
	wb <- createWorkbook()
	addWorksheet(wb, "Submission")
	addWorksheet(wb, "Full_Results")
	addWorksheet(wb, "Method")
	addWorksheet(wb, "Diagnostics")

	submission <- products_out %>%
		select(Rank, Product, PLC, Final_Forecast, Ref_DP, Ref_MKT, Ref_DS, Stat_Model, Ref_Ensemble)

	writeData(wb, "Submission", submission)
	writeData(wb, "Full_Results", products_out)

	method_note <- tibble(
		Item = c(
			"Backtest tuning window",
			"Robust holdout tuning quarters",
			"W_STAT guardrail",
			"Selected model set",
			"Statistical engine",
			"Efficiency approach",
			"Reference weighting",
			"Final blend",
			"Target"
		),
		Detail = c(
			paste(BACKTEST_QTRS, collapse = ", "),
			paste(TUNING_QTRS, collapse = ", "),
			sprintf("[%.2f, %.2f]", W_STAT_MIN, W_STAT_MAX),
			sprintf("%s (%s)", set_eval$selected_set, paste(selected_method_names, collapse = ", ")),
			"Weighted ensemble from tuned candidate set",
			"Weights learned using rolling-origin CV inverse-MAE scoring",
			"Global team weights from historical accuracy (DP/MKT/DS)",
			sprintf("%.0f%% Statistical + %.0f%% Reference", 100 * w_stat_tuned, 100 * w_ref_tuned),
			TARGET_QUARTER
		)
	)
	writeData(wb, "Method", method_note)

	# Diagnostics sheet: candidate set ranking, blend-tuning grid (top), and product backtest.
	diag_top <- tibble(
		Metric = c(
			"Selected_Model_Set",
			"Per_Product_Model_Set_Tuning",
			"Per_Product_Blend_Tuning",
			"W_STAT_Guardrail",
			"Tuning_Window",
			"Selected_Methods",
			"Backtest_Mean_Accuracy",
			"Backtest_Mean_Abs_Error",
			"Tuned_W_STAT",
			"Tuned_W_REF",
			"Backtest_Stat_Mean_Accuracy",
			"Backtest_Ref_Mean_Accuracy",
			"Backtest_Blended_Proxy_Mean_Accuracy"
		),
		Value = c(
			set_eval$selected_set,
			"Enabled",
			"Enabled",
			sprintf("[%.2f, %.2f]", W_STAT_MIN, W_STAT_MAX),
			paste(TUNING_QTRS, collapse = ", "),
			paste(selected_method_names, collapse = ", "),
			sprintf("%.4f", set_eval$summary$Mean_Accuracy[match(set_eval$selected_set, set_eval$summary$Model_Set)]),
			sprintf("%.2f", set_eval$summary$Mean_Abs_Error[match(set_eval$selected_set, set_eval$summary$Model_Set)]),
			sprintf("%.2f", w_stat_tuned),
			sprintf("%.2f", w_ref_tuned),
			sprintf("%.4f", validation_overall$Stat_Mean_Accuracy[1]),
			sprintf("%.4f", validation_overall$Ref_Mean_Accuracy[1]),
			sprintf("%.4f", validation_overall$Blended_Proxy_Mean_Accuracy[1])
		)
	)

	blend_top10 <- blend_tune$grid_scores %>%
		slice_head(n = 10)

	blend_product_top10 <- blend_tune_product$weights_by_rank %>%
		left_join(products %>% select(Rank, Product) %>% distinct(), by = "Rank") %>%
		arrange(desc(Product_Proxy_Mean_Accuracy)) %>%
		select(Rank, Product, W_STAT_Product, W_REF_Product, Product_Proxy_Mean_Accuracy) %>%
		slice_head(n = 10)

	model_set_by_product <- product_set_tune$best_by_rank %>%
		left_join(products %>% select(Rank, Product) %>% distinct(), by = "Rank") %>%
		select(Rank, Product, Model_Set, Mean_Accuracy, Mean_Abs_Error, N_Points) %>%
		arrange(desc(Mean_Accuracy), Mean_Abs_Error)

	product_backtest <- product_set_tune$selected_detail %>%
		left_join(products %>% select(Rank, Product) %>% distinct(), by = "Rank") %>%
		group_by(Rank, Product) %>%
		summarise(
			Backtest_Mean_Accuracy = mean(Stat_Accuracy, na.rm = TRUE),
			Backtest_MAE = mean(Abs_Error, na.rm = TRUE),
			.groups = "drop"
		) %>%
		arrange(desc(Backtest_Mean_Accuracy), Backtest_MAE)

	confidence_table <- product_confidence %>%
		left_join(products %>% select(Rank, Product) %>% distinct(), by = "Rank") %>%
		select(Rank, Product, Confidence_Score, Mean_Accuracy, SD_Accuracy, N_Points, Coverage_Factor, Stability_Factor, Mean_APE) %>%
		arrange(desc(Confidence_Score), desc(Mean_Accuracy))

	writeData(wb, "Diagnostics", diag_top, startRow = 1)
	writeData(wb, "Diagnostics", set_eval$summary, startRow = 10)
	row_q <- 10 + nrow(set_eval$summary) + 3
	writeData(wb, "Diagnostics", validation_by_quarter, startRow = row_q)
	row_blend <- row_q + nrow(validation_by_quarter) + 3
	writeData(wb, "Diagnostics", blend_top10, startRow = row_blend)
	row_blend_prod <- row_blend + nrow(blend_top10) + 4
	writeData(wb, "Diagnostics", blend_product_top10, startRow = row_blend_prod)
	row_model_set <- row_blend_prod + nrow(blend_product_top10) + 4
	writeData(wb, "Diagnostics", model_set_by_product, startRow = row_model_set)
	row_conf <- row_model_set + nrow(model_set_by_product) + 4
	writeData(wb, "Diagnostics", confidence_table, startRow = row_conf)
	row_prod <- row_conf + nrow(confidence_table) + 4
	writeData(wb, "Diagnostics", product_backtest, startRow = row_prod)

	header_style <- createStyle(textDecoration = "bold", fgFill = "#003B5C", fontColour = "white", halign = "center")
	addStyle(wb, "Submission", header_style, rows = 1, cols = 1:ncol(submission), gridExpand = TRUE)
	addStyle(wb, "Full_Results", header_style, rows = 1, cols = 1:ncol(products_out), gridExpand = TRUE)
	addStyle(wb, "Method", header_style, rows = 1, cols = 1:2, gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = 1, cols = 1:2, gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = 10, cols = 1:ncol(set_eval$summary), gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = row_q, cols = 1:ncol(validation_by_quarter), gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = row_blend, cols = 1:ncol(blend_top10), gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = row_blend_prod, cols = 1:ncol(blend_product_top10), gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = row_model_set, cols = 1:ncol(model_set_by_product), gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = row_conf, cols = 1:ncol(confidence_table), gridExpand = TRUE)
	addStyle(wb, "Diagnostics", header_style, rows = row_prod, cols = 1:ncol(product_backtest), gridExpand = TRUE)

	setColWidths(wb, "Submission", cols = 1:ncol(submission), widths = "auto")
	setColWidths(wb, "Full_Results", cols = 1:ncol(products_out), widths = "auto")
	setColWidths(wb, "Method", cols = 1:2, widths = c(28, 90))
	setColWidths(wb, "Diagnostics", cols = 1:6, widths = "auto")

	saveWorkbook(wb, output_file, overwrite = TRUE)

	cat(sprintf("Pipeline completed. Output written to: %s\n", output_file))

	products_out
}

# ---------------------------
# 6) Script entry point
# ---------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) >= 1) {
	DATA_FILE <- args[[1]]
}
if (length(args) >= 2) {
	OUTPUT_FILE <- args[[2]]
}

results <- run_pipeline(DATA_FILE, OUTPUT_FILE)

cat("\nTop 10 forecasted products:\n")
results %>%
	arrange(desc(Final_Forecast)) %>%
	select(Rank, Product, Final_Forecast) %>%
	slice_head(n = 10) %>%
	print(n = 10)