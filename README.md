# Upload_Fastq_to_GCP
Uploads and/or checks MD5Sum of Fastq files in local directory and GCP Bucket folder

## Description
upload_fastq_gcp.py is a Python-based automation utility designed for NGS and bioinformatics workflows to upload FASTQ files to Google Cloud Storage (GCP) and validate upload integrity using MD5 checksum comparison.

The script automatically detects all `.fastq.gz` files in the current working directory, uploads them to a specified GCP bucket/folder using `gsutil`, and compares local MD5 checksums against MD5 values retrieved from GCP. This helps ensure FASTQ files are uploaded without corruption or incomplete transfer.

The tool supports multiple operational modes including:
- FASTQ upload + MD5 validation with command, python3 upload_fastq_gcp.py <GCP_FOLDER_PATH>
- Validation-only mode without re-upload of FASTQ with command, python3 upload_fastq_gcp.py <GCP_FOLDER_PATH> --validate-only
- GCP folder MD5 reporting mode with command "--gcp-folder-md5", python3 upload_fastq_gcp.py <GCP_FOLDER_PATH> --gcp-folder-md5

