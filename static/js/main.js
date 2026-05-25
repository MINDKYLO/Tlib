// Auto-dismiss alerts after 4 seconds
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('.alert').forEach(function (alert) {
    setTimeout(function () {
      const bsAlert = bootstrap.Alert.getOrCreateInstance(alert);
      bsAlert.close();
    }, 4000);
  });
});
