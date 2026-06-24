# =============================================================
# StationDeck - Email Delivery Module
# src/emailer.py
# =============================================================
# Sends the monthly report package (PDF + DOCX + XLSX) to all
# configured recipients via Gmail SMTP with TLS encryption.
#
# Master function:
#   send_monthly_report(pdf_path, docx_path, xlsx_path,
#                       period_label, metrics, recipients=None)
#
# Called from main.py after the archive step.
# Non-fatal: logs errors but never crashes the pipeline.
# =============================================================

import smtplib
import logging
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime

from config.settings import (
    SMTP_HOST,
    SMTP_PORT,
    SMTP_USER,
    SMTP_PASS,
    REPORT_RECIPIENTS,
    STATION_NAME,
    STATION_LOCATION,
    REPORT_CURRENCY,
)

logger = logging.getLogger(__name__)


# =============================================================
# HELPER: Format numbers cleanly for the email body
# =============================================================

def _fmt(value, decimals=0):
    """Format a number with commas. Returns 'N/A' if value is missing."""
    try:
        return f"{float(value):,.{decimals}f}"
    except (TypeError, ValueError):
        return "N/A"


# =============================================================
# HELPER: Build the HTML email body
# =============================================================

def _build_html_body(period_label: str, metrics: dict,
                     station_name: str, station_location: str) -> str:
    """
    Returns a complete HTML email body with a key metrics summary table.
    Professional but simple — readable on mobile and desktop.
    """

    fuel_rev        = _fmt(metrics.get("total_fuel_revenue", 0))
    shop_rev        = _fmt(metrics.get("shop_sales_total", 0))
    total_sales     = _fmt(metrics.get("total_sales", 0))
    total_expenses  = _fmt(metrics.get("total_expenses", 0))
    cash_banked     = _fmt(metrics.get("total_cash_banked", 0))
    delta_status    = metrics.get("delta_status", "N/A")
    total_days      = metrics.get("total_days", "N/A")
    pms_vol         = _fmt(metrics.get("pms_volume_total", 0), 2)
    ago_vol         = _fmt(metrics.get("ago_volume_total", 0), 2)
    currency        = metrics.get("currency", REPORT_CURRENCY)

    pnl             = metrics.get("pnl", {})
    gross_income    = _fmt(pnl.get("gross_income", 0))
    net_profit      = _fmt(pnl.get("net_profit", 0))

    delta_colour = "#1E7E34" if delta_status.upper() == "SURPLUS" else "#C8102E"
    generated_at = datetime.now().strftime("%d %B %Y at %H:%M")

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0; padding:0; background:#f4f4f4; font-family: Arial, sans-serif;">

      <!-- Wrapper -->
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4; padding: 30px 0;">
        <tr><td align="center">

          <!-- Card -->
          <table width="620" cellpadding="0" cellspacing="0"
                 style="background:#ffffff; border-radius:8px;
                        box-shadow: 0 2px 8px rgba(0,0,0,0.1);">

            <!-- Header -->
            <tr>
              <td style="background:#1A1A2E; border-radius:8px 8px 0 0;
                         padding: 28px 32px; text-align:center;">
                <h1 style="margin:0; color:#ffffff; font-size:22px;
                            letter-spacing:1px;">StationDeck</h1>
                <p style="margin:6px 0 0; color:#C8102E; font-size:13px;
                           letter-spacing:2px; text-transform:uppercase;">
                  Monthly Report Delivery
                </p>
              </td>
            </tr>

            <!-- Intro -->
            <tr>
              <td style="padding: 28px 32px 16px;">
                <p style="margin:0; font-size:15px; color:#333;">Hello,</p>
                <p style="margin:12px 0 0; font-size:15px; color:#333; line-height:1.6;">
                  Your <strong>{period_label}</strong> operations report for
                  <strong>{station_name}</strong> ({station_location}) is ready.
                  Please find the full PDF, Word, and Excel reports attached.
                </p>
              </td>
            </tr>

            <!-- Metrics Table -->
            <tr>
              <td style="padding: 8px 32px 24px;">
                <p style="margin:0 0 12px; font-size:13px; font-weight:bold;
                           color:#1A1A2E; text-transform:uppercase;
                           letter-spacing:1px; border-bottom:2px solid #C8102E;
                           padding-bottom:6px;">
                  Key Highlights
                </p>
                <table width="100%" cellpadding="10" cellspacing="0"
                       style="border-collapse:collapse; font-size:14px;">

                  <tr style="background:#f0f0f0;">
                    <td style="color:#555; font-weight:bold; border-bottom:1px solid #ddd;">Metric</td>
                    <td style="color:#555; font-weight:bold; border-bottom:1px solid #ddd; text-align:right;">Value</td>
                  </tr>

                  <tr>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Reporting Period</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{period_label} &nbsp;({total_days} days)</td>
                  </tr>
                  <tr style="background:#fafafa;">
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Total Fuel Revenue</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {fuel_rev}</td>
                  </tr>
                  <tr>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">PMS Volume Sold</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{pms_vol} L</td>
                  </tr>
                  <tr style="background:#fafafa;">
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">AGO Volume Sold</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{ago_vol} L</td>
                  </tr>
                  <tr>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Shop Sales</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {shop_rev}</td>
                  </tr>
                  <tr style="background:#fafafa;">
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Total Station Sales</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {total_sales}</td>
                  </tr>
                  <tr>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Total Expenses</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {total_expenses}</td>
                  </tr>
                  <tr style="background:#fafafa;">
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Gross Income</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {gross_income}</td>
                  </tr>
                  <tr>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Net Profit</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {net_profit}</td>
                  </tr>
                  <tr style="background:#fafafa;">
                    <td style="color:#333; border-bottom:1px solid #f0f0f0;">Cash Banked</td>
                    <td style="color:#333; border-bottom:1px solid #f0f0f0; text-align:right;">{currency} {cash_banked}</td>
                  </tr>
                  <tr>
                    <td style="color:#333; font-weight:bold;">Cash Reconciliation</td>
                    <td style="font-weight:bold; text-align:right; color:{delta_colour};">{delta_status}</td>
                  </tr>

                </table>
              </td>
            </tr>

            <!-- Attachments note -->
            <tr>
              <td style="padding: 0 32px 24px;">
                <p style="margin:0; font-size:13px; color:#666; line-height:1.6;">
                  <strong>Attached files:</strong><br>
                  📄 PDF Report &nbsp;|&nbsp; 📝 Word Document &nbsp;|&nbsp; 📊 Excel Workbook
                </p>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="background:#f0f0f0; border-radius:0 0 8px 8px;
                         padding:16px 32px; text-align:center;">
                <p style="margin:0; font-size:12px; color:#888;">
                  Generated by <strong>StationDeck</strong> on {generated_at}
                  &nbsp;|&nbsp; {station_name}, {station_location}
                </p>
                <p style="margin:6px 0 0; font-size:11px; color:#aaa;">
                  This is an automated report. Do not reply to this email.
                </p>
              </td>
            </tr>

          </table>
        </td></tr>
      </table>

    </body>
    </html>
    """
    return html


# =============================================================
# HELPER: Attach a single file to the email message
# =============================================================

def _attach_file(msg: MIMEMultipart, file_path: str) -> None:
    """
    Opens a file from disk and attaches it to the email message object.
    Skips silently if the file does not exist.
    """
    path = Path(file_path)
    if not path.exists():
        logger.warning(f"Attachment not found, skipping: {file_path}")
        return

    with open(path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={path.name}")
    msg.attach(part)
    logger.info(f"Attached: {path.name}")


# =============================================================
# MASTER FUNCTION
# =============================================================

def send_monthly_report(
    pdf_path: str,
    docx_path: str,
    xlsx_path: str,
    period_label: str,
    metrics: dict,
    recipients: list = None       # ← NEW: passed from station_config
) -> bool:
    """
    Builds and sends the monthly report email with all 3 files attached.

    Parameters:
        pdf_path     : Full path to the generated PDF file
        docx_path    : Full path to the generated DOCX file
        xlsx_path    : Full path to the generated XLSX file
        period_label : Human-readable period e.g. "April 2026"
        metrics      : The full metrics dict from processor.py
        recipients   : List of email addresses. If None, falls back to .env

    Returns:
        True  if email sent successfully
        False if any error occurred (error is logged, pipeline continues)
    """

    # ── Resolve recipients: prefer passed-in list, fall back to .env ─────────
    if recipients is None or len(recipients) == 0:
        env_recipients = os.getenv("REPORT_RECIPIENTS", "")
        recipients = [r.strip() for r in env_recipients.split(",") if r.strip()]

    # ── Resolve station identity: prefer metrics dict, fall back to settings ──
    station_name     = metrics.get("station_name", STATION_NAME)
    station_location = metrics.get("location", STATION_LOCATION)

    # ── Guard: check configuration is present ────────────────────────────────
    if not SMTP_USER or not SMTP_PASS:
        logger.error("Email not sent: SMTP_USER or SMTP_PASS missing from .env")
        return False

    if not recipients:
        logger.error("Email not sent: no recipients configured")
        return False

    # ── Build the message ─────────────────────────────────────────────────────
    subject = f"StationDeck | {station_name} — {period_label} Report"

    msg = MIMEMultipart("mixed")
    msg["From"]    = f"StationDeck Reports <{SMTP_USER}>"
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject

    html_body = _build_html_body(
        period_label, metrics, station_name, station_location
    )
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    _attach_file(msg, pdf_path)
    _attach_file(msg, docx_path)
    _attach_file(msg, xlsx_path)

    # ── Send via Gmail SMTP with TLS ──────────────────────────────────────────
    try:
        logger.info(f"Connecting to SMTP server {SMTP_HOST}:{SMTP_PORT} ...")
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, recipients, msg.as_bytes())

        logger.info(f"Email sent successfully to: {', '.join(recipients)}")
        print(f"  ✅ Email delivered to: {', '.join(recipients)}")
        return True

    except smtplib.SMTPAuthenticationError:
        logger.error("SMTP authentication failed. Check SMTP_USER and SMTP_PASS in .env")
        print("  ❌ Email failed: authentication error. Check your App Password in .env")
        return False

    except smtplib.SMTPException as e:
        logger.error(f"SMTP error while sending email: {e}")
        print(f"  ❌ Email failed: {e}")
        return False

    except Exception as e:
        logger.error(f"Unexpected error in emailer: {e}")
        print(f"  ❌ Email failed (unexpected): {e}")
        return False