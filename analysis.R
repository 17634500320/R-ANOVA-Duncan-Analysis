# --- 1. 自动安装并加载所需的程序包 ---
# 定义需要的包列表
required_packages <- c("readxl", "agricolae", "dplyr", "tidyr", "writexl", "tibble")

# 检查每个包是否已安装，如果未安装则进行安装
for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg)
  }
}

# 加载所有需要的包
library(readxl)
library(agricolae)
library(dplyr)
library(tidyr)
library(writexl)
library(tibble) # tibble包用于行名处理

# --- 2. 配置区 ---
input_filename <- "data.xlsx"
output_filename <- "data_with_significance.xlsx"
alpha <- 0.05 # 显著性水平

# --- 3. 读取并处理数据 ---
cat("--- 成功读取原始数据 ---\n")
original_data <- read_excel(input_filename, col_names = FALSE)
# 为了后续操作方便，给列一个临时的、唯一的名称
original_colnames <- paste0("Group_", 1:ncol(original_data))
names(original_data) <- original_colnames
print(original_data)

long_data <- original_data %>%
  pivot_longer(cols = everything(), names_to = "Group", values_to = "Value") %>%
  na.omit()

# --- 4. 执行统计分析 ---
cat("\n--- 开始进行统计分析 ---\n")
aov_model <- aov(Value ~ as.factor(Group), data = long_data)
aov_summary <- summary(aov_model)
p_value_anova <- aov_summary[[1]][["Pr(>F)"]][1]

cat(paste("ANOVA P-value:", format(p_value_anova, scientific = FALSE, digits = 4), "\n"))

# 准备一个用于存储最终显著性字母的向量
final_letters <- character(ncol(original_data))
names(final_letters) <- original_colnames

if (p_value_anova >= alpha) {
  cat("\nANOVA结果不显著 (p >= 0.05)，无需进行事后检验。\n")
  final_letters[] <- "-"
  
} else {
  cat("\nANOVA结果显著 (p < 0.05)，开始进行Duncan's Test事后检验...\n")
  duncan_result <- duncan.test(aov_model, "as.factor(Group)", console = FALSE)
  
  cat("\nDuncan's Test 结果 (按均值排序):\n")
  print(duncan_result$groups)
  
  duncan_groups <- duncan_result$groups
  
  # 将duncan_groups中的字母，按照原始列的顺序填入final_letters向量
  # duncan_groups的行名就是组名 (Group_1, Group_2, ...)
  final_letters <- duncan_groups[original_colnames, "groups"]
}

# --- 5. 准备并输出最终文件 (修正后的稳健方法) ---
# 将原始数据（列名是临时的Group_1...）的值转换为字符
output_df <- as.data.frame(lapply(original_data, as.character))
# 将最终排好序的显著性字母向量添加为新的一行
output_df <- rbind(output_df, final_letters)

# 添加行名，以便在Excel中显示
output_df <- tibble::rownames_to_column(output_df, " ")
output_df[nrow(output_df), 1] <- "Significance"

# 将列名改回纯数字 1, 2, 3...
colnames(output_df)[-1] <- 1:ncol(original_data)

# 将结果写入新的Excel文件
write_xlsx(output_df, path = output_filename, col_names = TRUE, format_headers = TRUE)

cat(paste("\n--- 分析完成，结果已保存至 '", output_filename, "' ---\n"))
cat("显著性字母已按照原始数据 1, 2, 3... 的列顺序排列。\n")