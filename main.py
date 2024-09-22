import smtplib
import pandas as pd
import time
from time import strftime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl


# Configuration
hosts_credentials = [
    {
        "host": "mail.ut.ac.ir",
        "port": 587,
        "username": "kaveh.masoumi@ut.ac.ir",
        "password": "4URX7iE*ahQb@5",
    }
]
send_interval = 1800
from_email = "kaveh.masoumi@ut.ac.ir"
to_email = "kaveh.masoumi@ut.ac.ir"
subject = "Test Email"
body = "This is a test email sent by the monitoring script."


columns = [
    "Time",
    "Server",
    "Status",
    "Error",
    "RTT",
    "Latency",
    "Connection Time",
    "TLS Handshake Time",
    "SMTP Response Time",
    "Total Time",
]
df = pd.DataFrame(columns=columns)


def measure_smtp_transaction(host, port, username, password):
    total_start = time.time()
    status, error, rtt, latency, conn_time, tls_time, smtp_time = (
        "Failure",
        None,
        None,
        None,
        None,
        None,
        None,
    )
    try:
        conn_start = time.time()
        server = smtplib.SMTP(host, port, timeout=10)
        conn_end = time.time()
        conn_time = (conn_end - conn_start) * 1000

        tls_start = time.time()
        server.starttls(context=ssl.create_default_context())
        tls_end = time.time()
        tls_time = (tls_end - tls_start) * 1000

        smtp_start = time.time()
        server.login(username, password)
        message = MIMEMultipart()
        message["From"] = from_email
        message["To"] = to_email
        message["Subject"] = subject
        message.attach(MIMEText(body, "plain"))
        server.sendmail(from_email, to_email, message.as_string())
        smtp_end = time.time()
        smtp_time = (smtp_end - smtp_start) * 1000

        server.quit()

        status = "Success"
        total_end = time.time()
        total_time = (total_end - total_start) * 1000
        latency = total_time - (conn_time + tls_time + smtp_time)
        rtt = total_time

        return status, error, rtt, latency, conn_time, tls_time, smtp_time, total_time
    except Exception as e:
        error = str(e)
        total_end = time.time()
        total_time = (
            (total_end - total_start) * 1000 if "total_start" in locals() else None
        )
        return status, error, rtt, latency, conn_time, tls_time, smtp_time, total_time


def log_results(time_stamp, host_info, results):
    global df
    status, error, rtt, latency, conn_time, tls_time, smtp_time, total_time = results
    df = pd.concat(
        [
            df,
            pd.DataFrame(
                [
                    [
                        time_stamp,
                        f"{host_info['host']}:{host_info['port']}",
                        status,
                        error,
                        rtt,
                        latency,
                        conn_time,
                        tls_time,
                        smtp_time,
                        total_time,
                    ]
                ],
                columns=columns,
            ),
        ],
        ignore_index=True,
    )
    df.to_excel("results.xlsx", index=False)


def main():
    while True:
        for host_info in hosts_credentials:
            current_time = strftime("%Y-%m-%d %H:%M:%S")
            results = measure_smtp_transaction(
                host_info["host"],
                host_info["port"],
                host_info["username"],
                host_info["password"],
            )
            log_results(current_time, host_info, results)
        print(f"Waiting {send_interval} seconds until the next check...")
        time.sleep(send_interval)


if __name__ == "__main__":
    main()
