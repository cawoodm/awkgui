$1!~/^#/ {
   total[$3]++
}

END {
   for (i in total) print(i, total[i])
}