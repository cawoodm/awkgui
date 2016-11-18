$1!~/^#/ {
   total[$1]++
}

END {
   for (i in total) print(i, total[i])
}