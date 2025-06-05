{
  "_id": ObjectId("自动生成"),
  "basic_info": {
    "township": "大武乡",  // 乡镇
    "village": "边李村",   // 村委
    "risk_date": ISODate("2025-05-17T00:00:00Z"),  // 出险时间
    "growth_stage": "成熟期",  // 出险时间对应生长时期
    "loss_level": "轻"       // 报损程度
  },
  "sampling_info": {
    "farmer_name": "武小敏",  // 抽样农户名称
    "plot_name": "边李村地",   // 地块名称
    "average_spikes_per_mu": 31.9,  // 平均亩穗(万/亩)
    "average_grains_per_spike": 32.8,  // 平均穗粒数(粒/穗)
    "thousand_grain_weight": 41.5  // 平均千粒重(克)
  },
  "yield_data": {
    "current_yield": 369.1,  // 抽样地块平均产量(kg/亩)
    "historical_yield": 490.1,  // 当地前三年平均产量(kg/亩)
    "loss_percentage": 24.7  // 损失程度%
  },
  "statistics": {
    "avg_loss_same_level": 23.8  // 相同报损程度平均损失率%
  },
  "metadata": {
    "source_file": "大武乡散户损失程度情况表.xlsx",
    "import_date": ISODate("2023-03-15T00:00:00Z"),
    "calculated_fields": {
      "calculated_yield": true,
      "calculated_loss": true
    }
  }
}