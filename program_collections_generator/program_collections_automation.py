#!/usr/bin/env python3
"""
Program Collections Automation Script

This script automates the generation of program collection correction files.
It can be run manually or integrated into CI/CD pipelines.
"""

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class ProgramCollectionsGenerator:
    """Generates program collection correction files"""

    def __init__(self, base_path: str = None):
        self.base_path = Path(base_path) if base_path else Path.cwd()
        self.data_path = self.base_path / 'data'

        # URLs and configuration
        self.sheets_url = (
            "https://docs.google.com/spreadsheets/d/"
            "16ioasEqMoXuv2tgJs7xlMgD_sPLvrxs7Cp3rwz273YE/"
            "export?format=xlsx&gid=0"
        )
        self.staging_api = "https://api-staging.data.niaid.nih.gov/v1/query"
        self.prod_api = "https://api.data.niaid.nih.gov/v1/query"
        self.staging_metadata_api = "https://api-staging.data.niaid.nih.gov/v1/metadata"
        self.prod_metadata_api = "https://api.data.niaid.nih.gov/v1/metadata"

        # Find correction path
        self.correction_path = self._find_correction_path()

    def _find_correction_path(self) -> Path:
        """Find the nde-metadata-corrections directory"""
        possible_paths = [
            self.base_path.parent.parent / 'nde-metadata-corrections',
            self.base_path.parent / 'nde-metadata-corrections',
            Path.home() / 'nde-metadata-corrections',
        ]

        for path in possible_paths:
            if path.exists():
                logger.info(f"Found correction path: {path}")
                return path

        # Create temp directory if not found
        temp_path = Path('/tmp/nde-metadata-corrections')
        temp_path.mkdir(exist_ok=True)
        (temp_path / 'collections_corrections_staging').mkdir(exist_ok=True)
        (temp_path / 'collections_corrections_production').mkdir(exist_ok=True)
        logger.warning(f"Using temporary path: {temp_path}")
        return temp_path

    def load_config_files(self) -> Tuple[List[str], List[str], List[str], List[str]]:
        """Load configuration files"""
        # Load approved production programs
        approved_prod = []
        try:
            with open(self.data_path / 'approved_for_prod.txt', 'r') as f:
                approved_prod = [line.strip() for line in f if line.strip()]
        except FileNotFoundError:
            logger.warning("approved_for_prod.txt not found")

        # Load control transferred programs
        control_transferred = []
        try:
            with open(self.data_path / 'control_transferred.txt', 'r') as f:
                control_transferred = [line.strip()
                                       for line in f if line.strip()]
        except FileNotFoundError:
            logger.warning("control_transferred.txt not found")

        # Load NIH codes
        act_codes = []
        ic_codes = []
        try:
            df = pd.read_csv(self.data_path / 'NIH_activity_codes.csv')
            act_codes = [x.strip() for x in df['Activity Code'].unique()
                         if pd.notna(x)]
        except FileNotFoundError:
            logger.warning("NIH_activity_codes.csv not found")

        try:
            df = pd.read_csv(self.data_path /
                             'NIH_IC_codes.tsv', delimiter='\t')
            ic_codes = [x.strip() for x in df['Code'].unique() if pd.notna(x)]
        except FileNotFoundError:
            logger.warning("NIH_IC_codes.tsv not found")

        return approved_prod, control_transferred, act_codes, ic_codes

    def _get_google_sheets_credentials(self):
        """Get Google Sheets API credentials from environment or file"""
        # Try to get credentials from environment variable (for GitHub Actions)
        creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        if creds_json:
            try:
                credentials_info = json.loads(creds_json)
                credentials = service_account.Credentials.from_service_account_info(
                    credentials_info,
                    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                )
                return credentials
            except Exception as e:
                logger.warning(
                    f"Failed to load credentials from environment: {e}")

        # Try to load from service account file
        creds_file = self.data_path / 'service-account-key.json'
        if creds_file.exists():
            try:
                credentials = service_account.Credentials.from_service_account_file(
                    str(creds_file),
                    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                )
                return credentials
            except Exception as e:
                logger.warning(f"Failed to load credentials from file: {e}")

        return None

    def download_program_data(self) -> pd.DataFrame:
        """Download program metadata from Google Sheets using API authentication"""
        # Extract spreadsheet ID from the URL
        sheet_id = "16ioasEqMoXuv2tgJs7xlMgD_sPLvrxs7Cp3rwz273YE"

        try:
            logger.info("Downloading program metadata...")

            # Try authenticated Google Sheets API first
            credentials = self._get_google_sheets_credentials()
            if credentials:
                logger.info("Using Google Sheets API with service account")
                service = build('sheets', 'v4', credentials=credentials)

                # Get the data from the 'metadata' sheet
                range_name = 'metadata'  # Adjust if your sheet has a different name
                result = service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=range_name
                ).execute()

                values = result.get('values', [])
                if not values:
                    raise Exception("No data found in Google Sheet")

                # Convert to DataFrame
                headers = values[0]
                data = values[1:] if len(values) > 1 else []
                df = pd.DataFrame(data, columns=headers)

                logger.info(
                    f"Downloaded {len(df)} programs via Google Sheets API")
                return df

            else:
                logger.warning(
                    "No Google Sheets credentials found, trying direct download")
                # Fallback to direct download (will fail for private sheets)
                response = requests.get(self.sheets_url, timeout=60)
                response.raise_for_status()

                # Save and read Excel file
                temp_file = self.data_path / 'temp_collections.xlsx'
                with open(temp_file, 'wb') as f:
                    f.write(response.content)

                df = pd.read_excel(temp_file, sheet_name='metadata',
                                   engine='openpyxl')
                temp_file.unlink()

                logger.info(f"Downloaded {len(df)} programs")
                return df

        except Exception as e:
            logger.error(f"Download failed: {e}")
            # Fallback to local file
            local_file = self.data_path / 'Program Collections.xlsx'
            if local_file.exists():
                logger.warning("Using local file as fallback")
                return pd.read_excel(local_file, sheet_name='metadata',
                                     engine='openpyxl')
            raise

    def parse_array_text(self, text: str) -> List[str]:
        """Parse comma/pipe separated text"""
        if pd.isna(text) or text == 'not found':
            return []

        text = str(text).replace('*', '')
        if ',' in text:
            items = text.split(',')
        elif '|' in text:
            items = text.split('|')
        else:
            items = [text]

        return [x.strip() for x in items if x.strip()]

    def parse_grant_id(self, grant_id: str, act_codes: List[str],
                       ic_codes: List[str]) -> Dict[str, str]:
        """Parse grant ID into components"""
        result = {
            'icCode': 'not found',
            'serialNum': 'not found'
        }

        try:
            remaining = grant_id.strip()

            # Skip application type if present
            if remaining[0].isdigit() and not remaining[:2].isdigit():
                remaining = remaining[2:] if remaining[1] in '-_ ' else remaining[1:]

            # Parse activity code
            for code in act_codes:
                if remaining.startswith(code):
                    remaining = remaining[len(code):]
                    if remaining.startswith('-'):
                        remaining = remaining[1:]
                    break

            # Parse IC code
            for code in ic_codes:
                if remaining.startswith(code):
                    result['icCode'] = code
                    remaining = remaining[len(code):]
                    if remaining.startswith('-'):
                        remaining = remaining[1:]
                    break

            # Extract serial number (first 6 digits)
            serial = ''
            for char in remaining:
                if char.isdigit():
                    serial += char
                    if len(serial) >= 6:
                        break
                elif char == '-':
                    break

            if serial:
                result['serialNum'] = serial

        except Exception as e:
            logger.warning(f"Failed to parse {grant_id}: {e}")

        return result

    def search_records(self, grant_list: List[str],
                       environment: str = 'staging') -> List[str]:
        """Search for records matching grant IDs"""
        api_url = self.staging_api if environment == 'staging' else self.prod_api
        record_ids = set()

        for grant in grant_list:
            try:
                # Try wildcard search
                url = (f"{api_url}?q=funding.identifier:*{grant}*"
                       "&fields=_id,funding.identifier&size=500")

                response = requests.get(url, timeout=30)
                response.raise_for_status()
                data = response.json()

                if 'hits' in data:
                    for hit in data['hits']:
                        record_ids.add(hit['_id'])

            except Exception as e:
                logger.warning(f"Search failed for {grant}: {e}")

        return list(record_ids)

    def generate_files(self, df: pd.DataFrame, environment: str = 'both'):
        """Generate all correction files"""
        approved_prod, control_transferred, act_codes, ic_codes = (
            self.load_config_files()
        )

        # Filter valid programs
        valid_df = df[
            (~df['fundingIDList'].isna()) &
            (df['fundingIDList'] != 'not found') &
            (~df['niaidURL'].isna()) &
            (df['fileName'] != '--')
        ].copy()

        logger.info(f"Processing {len(valid_df)} valid programs")

        for _, row in valid_df.iterrows():
            try:
                filename = row['fileName']

                write_staging = environment in ['staging', 'both']
                write_production = (
                    filename in approved_prod and
                    environment in ['production', 'both']
                )

                # Skip entirely if neither environment applies
                if not write_staging and not write_production:
                    continue

                # Parse grant IDs once — used for both records search and
                # the fundingIdentifiers field in the correction JSON so the
                # hub can match new records on the fly by funding.
                search_grants = self._parse_search_grants(
                    row, act_codes, ic_codes
                )

                if write_staging:
                    self._create_metadata_file(
                        row, 'staging',
                        funding_identifiers=search_grants,
                    )
                    self._create_records_file(
                        row, 'staging', search_grants,
                        control_transferred,
                    )
                    logger.info(
                        f"Generated staging files for {filename}")

                if write_production:
                    self._create_metadata_file(
                        row, 'production',
                        funding_identifiers=search_grants,
                    )
                    self._create_records_file(
                        row, 'production', search_grants,
                        control_transferred,
                    )
                    logger.info(
                        f"Generated production files for {filename}")

            except Exception as e:
                logger.error(f"Error processing {filename}: {e}")

    def _parse_search_grants(self, row: pd.Series,
                             act_codes: List[str],
                             ic_codes: List[str]) -> List[str]:
        """Parse grant IDs from the row into searchable patterns.

        Returns a deduplicated list of patterns like 'AI123456' that can
        be used both for API record search and for on-the-fly funding
        matching in the hub upload wrapper.
        """
        grant_texts = self.parse_array_text(row['fundingIDList'])
        prior_grants = self.parse_array_text(
            row.get('PriorProjectGrantIDs', ''))

        search_grants = []
        seen = set()
        for grant_text in grant_texts + prior_grants:
            grant_text = grant_text.replace("*", "").strip()
            parsed = self.parse_grant_id(grant_text, act_codes, ic_codes)

            if (parsed['icCode'] != 'not found' and
                    parsed['serialNum'] != 'not found'):
                pattern = parsed['icCode'] + parsed['serialNum']
            else:
                pattern = grant_text

            if pattern and pattern not in seen:
                search_grants.append(pattern)
                seen.add(pattern)

        return search_grants

    def _create_metadata_file(self, row: pd.Series, environment: str,
                              funding_identifiers: List[str] = None):
        """Create metadata correction file.

        When *funding_identifiers* is provided, they are embedded in the
        JSON so the hub can match new records by funding on the fly.
        """
        filename = row['fileName']
        alt_names = self.parse_array_text(row.get('alternateName', ''))
        parent_orgs = self.parse_array_text(row.get('parentOrganization', ''))

        description = (f"{row['description']} For more information, "
                       f"visit the NIAID program page: {row['niaidURL']}")

        metadata = {
            "@type": "ResearchProject",
            "name": row["name"],
            "abstract": row["abstract"],
            "description": description,
            "alternateName": alt_names,
            "url": row["url"],
            "parentOrganization": parent_orgs
        }

        output_data = {"sourceOrganization": [metadata]}

        # Include funding identifiers for real-time matching in the hub.
        if funding_identifiers:
            output_data["fundingIdentifiers"] = funding_identifiers

        # Write file
        if environment == 'production':
            output_dir = (
                self.correction_path / 'collections_corrections_production'
            )
        else:
            output_dir = (
                self.correction_path / 'collections_corrections_staging'
            )

        output_dir.mkdir(exist_ok=True)
        output_file = output_dir / f'{filename}_correction.json'

        with open(output_file, 'w') as f:
            json.dump(output_data, f, indent=4)

    def _create_records_file(self, row: pd.Series, environment: str,
                             search_grants: List[str],
                             control_transferred: List[str]):
        """Create records file.

        *search_grants* is pre-parsed by ``_parse_search_grants``.
        """
        filename = row['fileName']

        # Search for records
        record_ids = self.search_records(search_grants, environment)

        # Handle control transferred programs
        if filename in control_transferred:
            prior_records = self._get_prior_records(filename, environment)
            record_ids = list(set(record_ids + prior_records))

        # Write records file
        if environment == 'production':
            output_dir = (
                self.correction_path / 'collections_corrections_production'
            )
        else:
            output_dir = (
                self.correction_path / 'collections_corrections_staging'
            )

        output_dir.mkdir(exist_ok=True)
        output_file = output_dir / f'{filename}_records.txt'

        with open(output_file, 'w') as f:
            for record_id in sorted(record_ids):
                if "immport" in record_id:
                    clean_id = record_id.replace("immport_", "")
                    f.write(
                        f'https://data.niaid.nih.gov/resources?id={clean_id}\n'
                    )
                else:
                    f.write(
                        'https://data.niaid.nih.gov/resources?id='
                        f'{record_id}\n'
                    )

    def _get_prior_records(self, filename: str, environment: str) -> List[str]:
        """Get prior records for programs with transferred control"""
        try:
            if environment == 'production':
                base_url = (
                    "https://raw.githubusercontent.com/"
                    "NIAID-Data-Ecosystem/nde-metadata-corrections/"
                    "refs/heads/main/collections_corrections_production/"
                )
            else:
                base_url = ("https://raw.githubusercontent.com/"
                            "NIAID-Data-Ecosystem/nde-metadata-corrections/"
                            "refs/heads/main/collections_corrections_staging/")

            url = f"{base_url}{filename}_records.txt"
            response = requests.get(url, timeout=30)
            response.raise_for_status()

            records = []
            for line in response.text.split('\n'):
                line = line.strip()
                if line and 'resources?id=' in line:
                    record_id = line.split('resources?id=')[-1]
                    records.append(record_id)

            return records

        except Exception as e:
            logger.warning(f"Could not get prior records for {filename}: {e}")
            return []

    def get_build_info(self, environment: str = 'staging') -> Dict[str, str]:
        """Get build information from the API metadata endpoint"""
        try:
            if environment == 'staging':
                url = self.staging_metadata_api
            else:
                url = self.prod_metadata_api

            response = requests.get(url, timeout=30)
            response.raise_for_status()

            data = response.json()
            build_info = {
                'build_date': data.get('build_date', ''),
                'build_version': data.get('build_version', ''),
                'biothing_type': data.get('biothing_type', '')
            }

            logger.info(f"Retrieved {environment} build info: "
                        f"version={build_info['build_version']}, "
                        f"date={build_info['build_date']}")
            return build_info

        except Exception as e:
            logger.error(f"Failed to get {environment} build info: {e}")
            return {}

    def check_for_new_build(self, environment: str = 'staging') -> bool:
        """Check if there's a new build since last run"""
        try:
            current_build = self.get_build_info(environment)
            if not current_build:
                return False

            # Store build info in the corrections repository so it persists
            # between GitHub Action runs (local data/ directory is ephemeral)
            build_file = (self.correction_path /
                          f'last_{environment}_build.json')

            if build_file.exists():
                with open(build_file, 'r') as f:
                    last_build = json.load(f)

                # Compare build versions/dates
                if (current_build.get('build_version') !=
                        last_build.get('build_version')):
                    logger.info(f"New {environment} build detected: "
                                f"{last_build.get('build_version')} -> "
                                f"{current_build.get('build_version')}")

                    # Update stored build info
                    with open(build_file, 'w') as f:
                        json.dump(current_build, f, indent=2)

                    return True
                else:
                    logger.info(f"No new {environment} build detected")
                    return False
            else:
                # First run - store current build info
                logger.info(f"First run - storing {environment} build info")
                with open(build_file, 'w') as f:
                    json.dump(current_build, f, indent=2)
                return True

        except Exception as e:
            logger.error(f"Error checking for new {environment} build: {e}")
            return False

    def should_update(self, environment: str = 'both',
                      force_update: bool = False) -> bool:
        """Determine if program collections should be updated"""
        if force_update:
            logger.info("Force update requested")
            return True

        # Check for new builds
        needs_update = False

        if environment in ['staging', 'both']:
            if self.check_for_new_build('staging'):
                logger.info("New staging build detected - update needed")
                needs_update = True

        if environment in ['production', 'both']:
            if self.check_for_new_build('production'):
                logger.info("New production build detected - update needed")
                needs_update = True

        return needs_update

    def run_automation(self, environment: str = 'both',
                       force_update: bool = False) -> bool:
        """Run the complete automation process with build checking"""
        try:
            logger.info(f"Starting program collections automation "
                        f"(environment: {environment})")

            # Check if update is needed based on build changes
            if not self.should_update(environment, force_update):
                logger.info("No update needed - no new builds detected")
                return True

            # Download and process data
            df = self.download_program_data()
            self.generate_files(df, environment)

            logger.info(
                "Program collections generation completed successfully!")
            return True

        except Exception as e:
            logger.error(f"Automation failed: {e}")
            return False


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Generate program collection correction files'
    )
    parser.add_argument(
        '--environment',
        choices=['staging', 'production', 'both'],
        default='both',
        help='Target environment'
    )
    parser.add_argument(
        '--base-path',
        help='Base path for the script'
    )
    parser.add_argument(
        '--force-update',
        action='store_true',
        help='Force update even if no new build detected'
    )

    args = parser.parse_args()

    try:
        generator = ProgramCollectionsGenerator(args.base_path)

        # Run automation with build monitoring
        success = generator.run_automation(args.environment, args.force_update)

        if success:
            logger.info(
                "Program collections automation completed successfully!")
        else:
            logger.error("Program collections automation failed!")
            sys.exit(1)

    except Exception as e:
        logger.error(f"Generation failed: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
