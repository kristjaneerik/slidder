import fire
import pickle
import os.path
import re
import mimetypes
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]  # read-write for all Slides (so that we can add images) and Drive (so that we can upload images)

img_file_pattern = "png|jpg|jpeg|gif|PNG|JPG|JPEG|GIF"
slidderpath_regex = re.compile(rf"path=([^\s]*\.(?:{img_file_pattern}))")


class GAPI(object):
    def __init__(self, client_secret_path="client_secret.json", token_path="token.pickle"):
        # first, set up credentials
        self.creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(token_path):
            with open(token_path, "rb") as token:
                self.creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_path, "wb") as token:
                pickle.dump(self.creds, token)

        # second, connect to services
        self.slides_service = build("slides", "v1", credentials=self.creds)
        self.drive_service = build("drive", "v3", credentials=self.creds)

    def get_presentation(self, presentation_id):
        return self.slides_service.presentations().get(presentationId=presentation_id).execute()

    def upload_image(self, localpath):
        file_hash = get_file_hash(localpath)
        file_metadata = {
            "name": file_hash,
            # "parents": ["appDataFolder"],
        }
        mimetype = mimetypes.guess_type(localpath)[0]
        media = MediaFileUpload(
            localpath, mimetype=mimetype, resumable=True,
        )
        file_obj = self.drive_service.files().create(
            body=file_metadata, media_body=media, fields="id",
        ).execute()
        obj_id = file_obj["id"]
        print(f"Uploaded {localpath} -> {file_hash} -> {obj_id}")
        return obj_id, self.get_uploaded_image_url(obj_id)

    def get_uploaded_image_url(self, drive_id):
        response = self.drive_service.files().get(
            fileId=drive_id, fields="webContentLink",
        ).execute()
        return response["webContentLink"]

    def remove_files(self, file_ids):
        if not file_ids:
            return
        for file_id in file_ids:
            self.drive_service.files().delete(fileId=file_id).execute()

    def make_public(self, file_ids):
        permission = {"type": "anyone", "role": "reader"}
        for file_id in file_ids:
            self.drive_service.permissions().create(
                fileId=file_id, body=permission, fields="id",
            ).execute()

    def _list_appdata_files(self):  # for debugging
        return self.drive_service.files().list(
            spaces="appDataFolder", fields="nextPageToken, files(id, name)", pageSize=100,
        ).execute()


def get_file_hash(localpath):
    return localpath  # TODO


def main(
    presentation_id,
    directory="./",
    debug=True,
):
    """
    """
    gapi = GAPI()
    presentation = gapi.get_presentation(presentation_id)
    slides = presentation.get("slides")

    if debug:
        print(f"The presentation contains {len(slides)} slides")
    requests = []
    uploaded_files = {}
    for slide in slides:
        images = [el for el in slide.get("pageElements", []) if "image" in el]
        slide_id = slide.get("objectId")
        for image in images:
            img_id = image.get("objectId")
            desc = image.get("description", "")
            # TODO need to re-set the description!
            url = image.get("image", {}).get("contentUrl")
            files = slidderpath_regex.findall(desc)
            if len(files) != 1:
                continue
            raw_fname = files[0]
            fname = os.path.join(directory, raw_fname)
            if debug:
                print(f"{img_id} on slide {slide_id}: {url} -> {raw_fname} -> {fname}")
            if not os.path.isfile(fname):
                print(
                    f"Tried to find {fname}, but didn't find it "
                    f"(looked for {raw_fname} in {directory})!"
                )
                continue
            if fname in uploaded_files:
                remote_url = uploaded_files[fname]["remote_url"]
            else:
                file_id, remote_url = gapi.upload_image(fname)
                uploaded_files[fname] = {"file_id": file_id, "remote_url": remote_url}
            requests.append({
                "replaceImage": {
                    "imageObjectId": img_id,
                    "imageReplaceMethod": "CENTER_CROP",
                    "url": remote_url,
                }
            })

    drive_file_ids = [f["file_id"] for f in uploaded_files.values()]
    if drive_file_ids and requests:
        # make all uploaded files world-readable-if-have-link so that we can import them
        gapi.make_public(drive_file_ids)

    response = None
    if requests:
        print("Have requests:")
        print(requests)
        response = gapi.slides_service.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests},
        ).execute()
        # print(response)

    gapi.remove_files(drive_file_ids)  # remove all uploaded files from the app directory


if __name__ == "__main__":
    fire.Fire(main)
