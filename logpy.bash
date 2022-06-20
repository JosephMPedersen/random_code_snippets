logpy () {
  python "$@" 2>&1 | tee "./logs/log_$(echo "${1%.*}")_$(date +%Y_%m_%d_%H%M%S).txt";
}
