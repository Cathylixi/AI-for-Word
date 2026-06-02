const express = require("express");

const router = express.Router();

// Reserved for future Mock TFL domain APIs.
router.get("/api/tfl/health", (req, res) => {
  res.json({ success: true, message: "TFL routes ready" });
});

module.exports = router;
