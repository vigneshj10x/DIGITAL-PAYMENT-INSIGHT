"""
PaySphere FinTech - Synthetic Data Generator
Generates realistic, anonymized operational datasets for Use Case 13.
Fixed seed: 42 for complete reproducibility.
"""

import os
import json
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

def generate_paysphere_data(seed=42, output_dir="data/raw"):
    np.random.seed(seed)
    os.makedirs(output_dir, exist_ok=True)
    
    # ---------------------------------------------------------
    # 1. CUSTOMERS (~5,000 rows)
    # ---------------------------------------------------------
    num_customers = 5000
    customer_ids = [f"CUST_{i+1:05d}" for i in range(num_customers)]
    
    segments = ["Retail", "Premium", "Student", "Small Business", "Corporate"]
    segment_weights = [0.45, 0.20, 0.18, 0.12, 0.05]
    
    regions = [
        "Tamil Nadu", "Karnataka", "Maharashtra", "Delhi", "Gujarat",
        "West Bengal", "Kerala", "Telangana", "Uttar Pradesh", "Rajasthan"
    ]
    region_weights = [0.18, 0.16, 0.18, 0.12, 0.08, 0.07, 0.06, 0.07, 0.05, 0.03]
    
    cust_segments = np.random.choice(segments, size=num_customers, p=segment_weights)
    cust_regions = np.random.choice(regions, size=num_customers, p=region_weights)
    cust_tenure = np.random.randint(1, 61, size=num_customers) # 1 to 60 months
    
    df_customers = pd.DataFrame({
        "customer_id": customer_ids,
        "customer_segment": cust_segments,
        "region": cust_regions,
        "tenure_months": cust_tenure
    })
    
    # ---------------------------------------------------------
    # 2. MERCHANTS (~500 rows)
    # ---------------------------------------------------------
    num_merchants = 500
    merchant_ids = [f"MERCH_{i+1:04d}" for i in range(num_merchants)]
    
    merchant_categories = [
        "Grocery", "Food & Dining", "Utilities", "Fuel", "Retail Shopping",
        "Electronics", "Travel", "Jewellery", "Gaming", "Crypto/Forex"
    ]
    category_weights = [0.22, 0.20, 0.12, 0.10, 0.14, 0.08, 0.06, 0.04, 0.02, 0.02]
    
    merch_cats = np.random.choice(merchant_categories, size=num_merchants, p=category_weights)
    merch_regions = np.random.choice(regions, size=num_merchants, p=region_weights)
    
    df_merchants = pd.DataFrame({
        "merchant_id": merchant_ids,
        "merchant_category": merch_cats,
        "region": merch_regions
    })
    merch_cat_lookup = dict(zip(df_merchants["merchant_id"], df_merchants["merchant_category"]))
    merch_reg_lookup = dict(zip(df_merchants["merchant_id"], df_merchants["region"]))
    cust_reg_lookup = dict(zip(df_customers["customer_id"], df_customers["region"]))
    cust_seg_lookup = dict(zip(df_customers["customer_id"], df_customers["customer_segment"]))
    
    # ---------------------------------------------------------
    # 3. DEVICES (~8,000 rows)
    # ---------------------------------------------------------
    num_devices = 8000
    device_ids = [f"DEV_{i+1:05d}" for i in range(num_devices)]
    dev_customers = np.random.choice(customer_ids, size=num_devices)
    
    device_types = ["Android", "iOS", "Web Browser", "POS Terminal"]
    dev_type_weights = [0.60, 0.25, 0.12, 0.03]
    dev_types = np.random.choice(device_types, size=num_devices, p=dev_type_weights).astype(object)
    
    # New device flag: ~18% new devices
    new_device_flags = np.random.choice([0, 1], size=num_devices, p=[0.82, 0.18]).astype(object)
    
    # Quality issue 5: Incomplete device history (missing device_type / new_device_flag for ~4% records)
    incomplete_device_idx = np.random.choice(num_devices, size=int(num_devices * 0.04), replace=False)
    dev_types[incomplete_device_idx] = np.nan
    new_device_flags[incomplete_device_idx] = np.nan
    
    df_devices = pd.DataFrame({
        "device_id": device_ids,
        "customer_id": dev_customers,
        "device_type": dev_types,
        "new_device_flag": new_device_flags
    })
    
    # Create customer-to-devices mapping
    cust_to_devices = {}
    for dev_id, c_id in zip(device_ids, dev_customers):
        cust_to_devices.setdefault(c_id, []).append(dev_id)
        
    # ---------------------------------------------------------
    # 4. TRANSACTIONS (~50,000 rows base)
    # ---------------------------------------------------------
    num_transactions = 50000
    txn_ids = [f"TXN_{i+1:06d}" for i in range(num_transactions)]
    
    # Timestamps spanning ~6 months (2024-01-01 to 2024-06-30)
    start_date = datetime(2024, 1, 1, 0, 0, 0)
    total_seconds = int((datetime(2024, 6, 30, 23, 59, 59) - start_date).total_seconds())
    
    # Hourly probability weights (diurnal cycle: peak 10 AM to 9 PM, low 1 AM to 4 AM)
    hourly_probs = np.array([
        0.012, 0.005, 0.003, 0.003, 0.006, 0.015, # 00:00 - 05:00
        0.030, 0.045, 0.060, 0.075, 0.080, 0.085, # 06:00 - 11:00
        0.085, 0.075, 0.070, 0.072, 0.078, 0.082, # 12:00 - 17:00
        0.080, 0.065, 0.050, 0.035, 0.025, 0.019  # 18:00 - 23:00
    ])
    hourly_probs = hourly_probs / hourly_probs.sum()
    
    # Sample days and hours
    day_offsets = np.random.randint(0, 182, size=num_transactions)
    hours = np.random.choice(24, size=num_transactions, p=hourly_probs)
    minutes = np.random.randint(0, 60, size=num_transactions)
    seconds = np.random.randint(0, 60, size=num_transactions)
    
    timestamps = [
        start_date + timedelta(days=int(d), hours=int(h), minutes=int(m), seconds=int(s))
        for d, h, m, s in zip(day_offsets, hours, minutes, seconds)
    ]
    # Sort chronologically initially
    timestamps.sort()
    
    # Assign customers with frequency variation (Pareto-like customer activity)
    cust_probs = np.random.pareto(a=1.8, size=num_customers) + 0.1
    cust_probs = cust_probs / cust_probs.sum()
    txn_customers = np.random.choice(customer_ids, size=num_transactions, p=cust_probs)
    
    # Merchants
    merch_probs = np.random.pareto(a=1.5, size=num_merchants) + 0.1
    merch_probs = merch_probs / merch_probs.sum()
    txn_merchants = np.random.choice(merchant_ids, size=num_transactions, p=merch_probs)
    
    # Regional spike injection: Maharashtra and Delhi have a sharp surge in Week 14 (mid-April festival/promo)
    # Re-assign some merchants during that week to Maharashtra/Delhi to create realistic localized spikes
    for idx in range(num_transactions):
        ts = timestamps[idx]
        if ts.isocalendar()[1] == 15 and np.random.rand() < 0.25:
            # Force high-volume merchant in Maharashtra or Delhi
            cand_merch = [m for m in merchant_ids[:50] if merch_reg_lookup.get(m) in ["Maharashtra", "Delhi"]]
            if cand_merch:
                txn_merchants[idx] = np.random.choice(cand_merch)
                
    # Devices: assign based on customer
    txn_devices = []
    # Quality issue 5 part b: Transactions with device_id not found in devices.csv (~2.5%)
    unregistered_device_count = int(num_transactions * 0.025)
    unregistered_indices = set(np.random.choice(num_transactions, size=unregistered_device_count, replace=False))
    
    dev_lookup_dict = {}
    for row in df_devices.itertuples():
        dev_lookup_dict[row.device_id] = (row.device_type, row.new_device_flag)
        
    for i, c_id in enumerate(txn_customers):
        if i in unregistered_indices:
            txn_devices.append(f"DEV_UNREG_{np.random.randint(90000, 99999)}")
        else:
            cust_devs = cust_to_devices.get(c_id, [])
            if cust_devs:
                txn_devices.append(np.random.choice(cust_devs))
            else:
                # Fallback to random known device
                txn_devices.append(np.random.choice(device_ids))
                
    # Merchant Category from lookup
    txn_categories = [merch_cat_lookup.get(m, "Unknown") for m in txn_merchants]
    
    # Inconsistent Channel Labels (Quality issue 3)
    channel_variants = {
        "UPI": ["UPI", "upi", "Upi "],
        "Mobile App": ["Mobile App", "mobile_app"],
        "Web": ["Web", "web"],
        "POS": ["POS", "pos"],
        "ATM": ["ATM"]
    }
    channel_base_probs = [0.55, 0.20, 0.12, 0.09, 0.04]
    base_channels = np.random.choice(list(channel_variants.keys()), size=num_transactions, p=channel_base_probs)
    txn_channels = [np.random.choice(channel_variants[bc]) for bc in base_channels]
    
    # Transaction Amounts based on Customer Segment (Log-normal distributions)
    # Segment parameters: (meanlog, sigmalog)
    seg_amount_params = {
        "Student": (5.9, 0.65),         # Median ~ ₹365, 95th ~ ₹1,060
        "Retail": (7.4, 0.75),          # Median ~ ₹1,635, 95th ~ ₹5,800
        "Premium": (9.0, 0.85),         # Median ~ ₹8,100, 95th ~ ₹32,000
        "Small Business": (9.8, 0.90),  # Median ~ ₹18,000, 95th ~ ₹80,000
        "Corporate": (11.0, 1.05)       # Median ~ ₹60,000, 95th ~ ₹340,000
    }
    
    txn_amounts = []
    for c_id, dev_id, cat in zip(txn_customers, txn_devices, txn_categories):
        seg = cust_seg_lookup[c_id]
        meanlog, sigmalog = seg_amount_params[seg]
        amt = np.random.lognormal(meanlog, sigmalog)
        
        # New devices or high-risk categories slightly higher
        dev_info = dev_lookup_dict.get(dev_id, (None, None))
        if dev_info[1] == 1 or "UNREG" in dev_id:
            amt *= np.random.uniform(1.15, 1.45)
        if cat in ["Jewellery", "Electronics", "Crypto/Forex", "Travel"]:
            amt *= np.random.uniform(1.20, 1.60)
            
        txn_amounts.append(round(amt, 2))
    txn_amounts = np.array(txn_amounts)
    
    # Quality Issue 4: Extreme amounts and zero/negative amounts
    # Handful of very large outliers (~35 txns)
    large_outlier_idx = np.random.choice(num_transactions, size=35, replace=False)
    for idx in large_outlier_idx:
        txn_amounts[idx] = round(np.random.uniform(650000.0, 3200000.0), 2)
        
    # Zero and negative amounts (~30 txns)
    invalid_amount_idx = np.random.choice([i for i in range(num_transactions) if i not in large_outlier_idx], size=30, replace=False)
    for i, idx in enumerate(invalid_amount_idx):
        if i % 3 == 0:
            txn_amounts[idx] = 0.0
        else:
            txn_amounts[idx] = round(-1.0 * np.random.uniform(50.0, 2500.0), 2)
            
    # Quality Issue 1: Missing merchant categories (~4.2% missing)
    missing_cat_count = int(num_transactions * 0.042)
    missing_cat_idx = np.random.choice(num_transactions, size=missing_cat_count, replace=False)
    txn_categories = np.array(txn_categories, dtype=object)
    txn_categories[missing_cat_idx] = np.nan
    
    # Assemble Base Transactions DataFrame
    df_transactions = pd.DataFrame({
        "transaction_id": txn_ids,
        "customer_id": txn_customers,
        "device_id": txn_devices,
        "merchant_id": txn_merchants,
        "timestamp": [ts.strftime("%Y-%m-%d %H:%M:%S") for ts in timestamps],
        "amount": txn_amounts,
        "merchant_category": txn_categories,
        "channel": txn_channels
    })
    
    # Quality Issue 2: Duplicate transaction IDs (~1.4% duplicates, including exact duplicate rows)
    num_duplicates = int(num_transactions * 0.014) # ~700 duplicates
    dup_sample_indices = np.random.choice(num_transactions, size=num_duplicates, replace=False)
    
    exact_dup_count = num_duplicates // 2 # ~350 exact duplicates
    modified_dup_count = num_duplicates - exact_dup_count # ~350 with slight change (e.g. timestamp or amount)
    
    exact_dups = df_transactions.iloc[dup_sample_indices[:exact_dup_count]].copy()
    modified_dups = df_transactions.iloc[dup_sample_indices[exact_dup_count:]].copy()
    # Modify timestamp slightly for modified duplicates to simulate re-attempt or duplicate ID collision
    modified_dups["amount"] = modified_dups["amount"].apply(lambda a: round(a * 1.05, 2) if a > 0 else a)
    
    df_transactions_with_dups = pd.concat([df_transactions, exact_dups, modified_dups], ignore_index=True)
    # Shuffle slightly to intersperse duplicates
    df_transactions_with_dups = df_transactions_with_dups.sample(frac=1.0, random_state=seed).reset_index(drop=True)
    
    # ---------------------------------------------------------
    # 5. ALERTS (~3,800 rows on distinct transactions, overall rate ~7.5%)
    # ---------------------------------------------------------
    # Calculate anomaly score/probability for each transaction to assign alerts realistically
    alert_scores = []
    
    # Customer velocity calculation for alert generation
    # Sort distinct transactions by customer and timestamp
    df_distinct = df_transactions.copy()
    df_distinct["dt"] = pd.to_datetime(df_distinct["timestamp"])
    df_distinct = df_distinct.sort_values(["customer_id", "dt"])
    df_distinct["time_diff_min"] = df_distinct.groupby("customer_id")["dt"].diff().dt.total_seconds() / 60.0
    velocity_burst_ids = set(df_distinct[df_distinct["time_diff_min"] < 15]["transaction_id"])
    
    high_risk_categories = {"Electronics", "Gaming", "Crypto/Forex", "Jewellery", "Travel"}
    
    alert_reasons = []
    alert_candidates = []
    
    for row in df_transactions.itertuples():
        score = 0.0
        reasons = []
        
        # 1. Hour risk: 1-4 AM
        hr = int(row.timestamp[11:13])
        if hr in [1, 2, 3, 4]:
            score += 3.2
            reasons.append("Unusual Hour")
            
        # 2. Device risk: new device or unregistered
        dev_info = dev_lookup_dict.get(row.device_id, (None, None))
        is_new_dev = (dev_info[1] == 1) or ("UNREG" in str(row.device_id))
        if is_new_dev:
            score += 2.8
            reasons.append("New Device")
            
        # 3. High amount / extreme amount
        c_seg = cust_seg_lookup.get(row.customer_id, "Retail")
        amt = row.amount
        if amt > 250000:
            score += 4.5
            reasons.append("High Amount")
        elif c_seg == "Student" and amt > 3500:
            score += 2.5
            reasons.append("High Amount")
        elif c_seg == "Retail" and amt > 20000:
            score += 2.2
            reasons.append("High Amount")
        elif c_seg == "Premium" and amt > 90000:
            score += 2.2
            reasons.append("High Amount")
            
        # 4. Merchant Category risk
        cat = merch_cat_lookup.get(row.merchant_id, "Unknown")
        if cat in high_risk_categories:
            score += 1.8
            reasons.append("Unusual Merchant")
            
        # 5. Velocity burst
        if row.transaction_id in velocity_burst_ids:
            score += 2.6
            reasons.append("Velocity Burst")
            
        # 6. Region Anomaly: customer region != merchant region
        c_reg = cust_reg_lookup.get(row.customer_id)
        m_reg = merch_reg_lookup.get(row.merchant_id)
        if c_reg and m_reg and c_reg != m_reg:
            score += 1.4
            reasons.append("Region Anomaly")
            
        # Background baseline noise
        score += np.random.exponential(scale=0.8)
        
        # Primary reason
        primary_reason = reasons[0] if reasons else "High Amount"
        alert_scores.append(score)
        alert_reasons.append(primary_reason)
        
    df_transactions_eval = df_transactions.copy()
    df_transactions_eval["alert_score"] = alert_scores
    df_transactions_eval["reason"] = alert_reasons
    
    # Target ~3,800 alerts out of 50,000 unique transactions (~7.6% alert rate)
    target_alerts = 3850
    alerted_df = df_transactions_eval.nlargest(target_alerts, "alert_score")
    
    # Review statuses with realistic distribution
    review_statuses = ["Pending", "Under Review", "Confirmed Suspicious", "False Positive", "Escalated"]
    status_weights = [0.20, 0.25, 0.24, 0.21, 0.10]
    
    alert_records = []
    for i, row in enumerate(alerted_df.itertuples()):
        alert_id = f"ALT_{i+1:05d}"
        status = np.random.choice(review_statuses, p=status_weights)
        txn_dt = datetime.strptime(row.timestamp, "%Y-%m-%d %H:%M:%S")
        alert_dt = txn_dt + timedelta(minutes=int(np.random.randint(1, 45)), seconds=int(np.random.randint(0, 60)))
        alert_records.append({
            "alert_id": alert_id,
            "transaction_id": row.transaction_id,
            "review_status": status,
            "reason_category": row.reason,
            "alert_time": alert_dt.strftime("%Y-%m-%d %H:%M:%S")
        })
        
    df_alerts = pd.DataFrame(alert_records)
    
    # Save all datasets
    df_transactions_with_dups.to_csv(os.path.join(output_dir, "transactions.csv"), index=False)
    df_customers.to_csv(os.path.join(output_dir, "customers.csv"), index=False)
    df_devices.to_csv(os.path.join(output_dir, "devices.csv"), index=False)
    df_merchants.to_csv(os.path.join(output_dir, "merchants.csv"), index=False)
    df_alerts.to_csv(os.path.join(output_dir, "alerts.csv"), index=False)
    
    # Summary of Injected Quality Issues
    quality_summary = {
        "dataset_shapes": {
            "transactions_raw": df_transactions_with_dups.shape,
            "customers": df_customers.shape,
            "devices": df_devices.shape,
            "merchants": df_merchants.shape,
            "alerts": df_alerts.shape
        },
        "quality_issues": {
            "missing_merchant_categories": {
                "count": int(df_transactions_with_dups["merchant_category"].isna().sum()),
                "percentage": round(float(df_transactions_with_dups["merchant_category"].isna().mean() * 100), 2)
            },
            "duplicate_transactions": {
                "total_rows": len(df_transactions_with_dups),
                "unique_transaction_ids": int(df_transactions_with_dups["transaction_id"].nunique()),
                "duplicate_ids_count": int(df_transactions_with_dups["transaction_id"].duplicated().sum()),
                "exact_duplicate_rows": int(df_transactions_with_dups.duplicated().sum())
            },
            "inconsistent_channel_labels": {
                "unique_labels": list(df_transactions_with_dups["channel"].value_counts().to_dict().keys()),
                "counts": {k: int(v) for k, v in df_transactions_with_dups["channel"].value_counts().to_dict().items()}
            },
            "extreme_and_invalid_amounts": {
                "large_outliers_over_500k": int((df_transactions_with_dups["amount"] > 500000).sum()),
                "max_amount": float(df_transactions_with_dups["amount"].max()),
                "zero_amounts": int((df_transactions_with_dups["amount"] == 0).sum()),
                "negative_amounts": int((df_transactions_with_dups["amount"] < 0).sum()),
                "min_amount": float(df_transactions_with_dups["amount"].min())
            },
            "incomplete_device_history": {
                "missing_device_type_in_devices": int(df_devices["device_type"].isna().sum()),
                "missing_new_device_flag_in_devices": int(df_devices["new_device_flag"].isna().sum()),
                "transactions_with_unregistered_devices": int((~df_transactions_with_dups["device_id"].isin(df_devices["device_id"])).sum())
            },
            "alert_distribution": {
                "total_alerts": len(df_alerts),
                "unique_alerted_transactions": int(df_alerts["transaction_id"].nunique()),
                "alert_rate_pct": round(float(len(df_alerts) / df_transactions_with_dups["transaction_id"].nunique() * 100), 2),
                "review_status_breakdown": {k: int(v) for k, v in df_alerts["review_status"].value_counts().to_dict().items()},
                "reason_category_breakdown": {k: int(v) for k, v in df_alerts["reason_category"].value_counts().to_dict().items()}
            }
        }
    }
    
    with open(os.path.join(output_dir, "data_quality_injection_summary.json"), "w") as f:
        json.dump(quality_summary, f, indent=4)
        
    print(f"Data generation complete! Saved to {output_dir}/")
    print(json.dumps(quality_summary["quality_issues"], indent=2))
    return quality_summary

if __name__ == "__main__":
    generate_paysphere_data()
