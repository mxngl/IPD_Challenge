import argparse
import os
import sys
from urllib.parse import unquote

import requests

APS_BASE_URL = "https://developer.api.autodesk.com"
TOKEN_URL = APS_BASE_URL + "/authentication/v2/token"
DATA_BASE_URL = APS_BASE_URL + "/data/v1"
OSS_BASE_URL = APS_BASE_URL + "/oss/v2"


class ApsClient(object):
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.session = requests.Session()
        self._token = None

    def authenticate(self):
        response = self.session.post(
            TOKEN_URL,
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "data:read bucket:read",
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()
        self._token = payload["access_token"]
        self.session.headers.update({"Authorization": "Bearer {0}".format(self._token)})
        return payload

    def get_version_details(self, project_id, version_id):
        url = "{0}/projects/{1}/versions/{2}".format(DATA_BASE_URL, project_id, version_id)
        response = self.session.get(url, timeout=30)
        response.raise_for_status()
        return response.json()

    def create_signed_download_url(self, bucket_key, object_name):
        encoded_object = requests.utils.quote(object_name, safe="")
        url = "{0}/buckets/{1}/objects/{2}/signeds3download".format(
            OSS_BASE_URL,
            bucket_key,
            encoded_object,
        )
        response = self.session.get(url, timeout=30)
        response.raise_for_status()
        return response.json()

    def download_file(self, signed_url, destination):
        with self.session.get(signed_url, stream=True, timeout=120) as response:
            response.raise_for_status()
            with open(destination, "wb") as file_handle:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        file_handle.write(chunk)


def parse_storage_urn(storage_urn):
    prefix = "urn:adsk.objects:os.object:"
    if not storage_urn.startswith(prefix):
        raise ValueError("Unsupported storage URN: {0}".format(storage_urn))

    bucket_and_object = storage_urn[len(prefix):]
    bucket_key, object_name = bucket_and_object.split("/", 1)
    return bucket_key, unquote(object_name)


def extract_storage_urn(version_payload):
    storage = version_payload.get("data", {}).get("relationships", {}).get("storage", {}).get("data")
    if not storage or "id" not in storage:
        raise ValueError(
            "This version payload does not include a storage relationship. "
            "Confirm the version id belongs to an ACC/BIM 360 file version."
        )
    return storage["id"]


def build_arg_parser():
    parser = argparse.ArgumentParser(
        description="Download a Revit file version from Autodesk Platform Services (ACC/BIM 360)."
    )
    parser.add_argument("--project-id", required=True, help="ACC/BIM 360 project id")
    parser.add_argument("--version-id", required=True, help="Version id from the Data Management API")
    parser.add_argument("--output", required=True, help="Destination .rvt file path")
    parser.add_argument("--client-id", default=os.getenv("APS_CLIENT_ID"), help="APS client id")
    parser.add_argument("--client-secret", default=os.getenv("APS_CLIENT_SECRET"), help="APS client secret")
    return parser


def main():
    parser = build_arg_parser()
    args = parser.parse_args()

    if not args.client_id or not args.client_secret:
        parser.error("Provide --client-id/--client-secret or set APS_CLIENT_ID and APS_CLIENT_SECRET.")

    client = ApsClient(args.client_id, args.client_secret)
    client.authenticate()

    version_payload = client.get_version_details(args.project_id, args.version_id)
    storage_urn = extract_storage_urn(version_payload)
    bucket_key, object_name = parse_storage_urn(storage_urn)
    signed_download = client.create_signed_download_url(bucket_key, object_name)
    signed_url = signed_download.get("url")

    if not signed_url:
        raise RuntimeError("APS did not return a signed download URL.")

    output_dir = os.path.dirname(os.path.abspath(args.output))
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    client.download_file(signed_url, args.output)
    sys.stdout.write("Downloaded {0}\n".format(args.output))


if __name__ == "__main__":
    main()
