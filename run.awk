function formatNum(num) {
  gsub("\"", "", num)
  gsub(",", "", num)
  return num+0
}
BEGIN {
  FS=","
  OFS=","
  FPAT = "([^,]+)|(\"[^\"]+\")"
}
FNR > 1 { 
  # 文件1中，$1为『人名』，$2为『数值』
  if (length($1)!=0) {
    if (NR==FNR) {
      if ($2~/-/) {
        obj[$1]=0
      } else {
        # gsub(/ |\t/, "", $2)
        obj[$1]=formatNum($2)
      }
    } else if($2 in obj) {
      # 文件2中，$2为『人名』，$3为『数值』
      result[$2]=obj[$2]-formatNum($3)
      delete obj[$2]
    } else {
      notInSheetOne[$2]=0-formatNum($3)
    }
  }
}
END {
  print "职员姓名","结果"
  for (key in result) {
    printf "%s, %.2f\n", key,result[key]
  }
  for (key in obj) {
    printf "%s, %.2f\n", key,obj[key]
  }
  for (key in notInSheetOne) {
    printf "%s, %.2f\n", key,notInSheetOne[key]
  }
}
