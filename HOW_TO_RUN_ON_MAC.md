# Run the PK vs no-PK experiment on your MacBook with Docker

MySQL runs inside a Docker container on your laptop, so you don't install MySQL itself.
The whole thing takes about 20 minutes (3 full runs + 1 verification run).

## 1. Install Docker Desktop (one time)

1. Download **Docker Desktop for Mac** from https://www.docker.com/products/docker-desktop/
   - pick **Apple Silicon** (M1/M2/M3/M4) or **Intel chip**, whichever your Mac has
     (Apple menu > About This Mac shows it).
2. Open Docker Desktop and wait until it says **Engine running**.
3. Check in **Terminal**:

```bash
docker --version
docker compose version
```

## 2. Run the experiment (one command)

Unzip `pk_experiment_scripts.zip`, then in Terminal:

```bash
cd ~/Downloads/pk_experiment_scripts     # the unzipped folder
docker compose up --abort-on-container-exit
```

What happens:

1. Docker downloads MySQL 8.0 and Python (first time only, a few minutes).
2. MySQL starts; the experiment waits until it is ready.
3. You see each timing printed as it goes (`pk-experiment | SELECT id= ...`).
4. At the end it writes three files **into the same folder**:
   **pk_experiment_results.xlsx**, results.json and verification.json.
5. Everything stops by itself when it's finished.

The Excel file has the same layout as the one I sent, with your laptop's numbers.
Checklist item 1 will show **Done** ("It ran on your laptop, with MySQL in a Docker container").

When you are finished, clean up with:

```bash
docker compose down
```

## 3. Optional: screenshots in MySQL Workbench or the terminal

Start only MySQL and leave it running:

```bash
docker compose up -d mysql
```

- **Terminal:** `docker exec -it pk-mysql mysql -uroot -proot pk_experiment`
  then type queries such as `SELECT * FROM data_nopk WHERE id = 1000000;`
  (MySQL prints the time, e.g. `1 row in set (0.49 sec)`).
- **MySQL Workbench:** new connection, host `127.0.0.1`, port `3306`, user `root`, password `root`.
- **Pure-SQL version of the whole experiment:**
  `docker exec -i pk-mysql mysql -uroot -proot < pk_experiment.sql`
  (`SHOW PROFILES` in the output lists every query's duration in seconds).

Stop it afterwards with `docker compose down`.

## If something goes wrong

| Problem | Fix |
|---|---|
| `Cannot connect to the Docker daemon` | Open Docker Desktop and wait for "Engine running". |
| `port 3306 is already allocated` / `address already in use` | Another MySQL is running on your Mac. Stop it (`brew services stop mysql`) or change the line `"3306:3306"` in docker-compose.yml to `"3307:3306"`. |
| The experiment stops with `Killed` | Docker needs more memory: Docker Desktop > Settings > Resources > Memory, set 4 GB or more. |
| You want a quick test first | In docker-compose.yml add `--runs 1` at the end of the `python -u pk_experiment.py ...` line (about 7 minutes). |

## What to expect

- SELECT **with PK**: well under 1 ms for every row id (index lookup, EXPLAIN `type = const`, 1 row read).
- SELECT **without PK**: hundreds of ms, about the same for row 10 and row 1,000,000
  (full table scan, EXPLAIN `type = ALL`, all 1,000,000 rows read).
- INSERT: both tables are close; the PK barely costs anything because ids are increasing.

Docker adds a little overhead on a Mac, so absolute times will differ from my reference run.
Compare the **ratio** between the two tables, not the raw milliseconds.
