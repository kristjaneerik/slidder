import fire
import pickle
import os.path
import re
import mimetypes
import hashlib
import urllib
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]  # read-write for all Slides (so that we can add images) and Drive (so that we can upload images)

img_file_regex = re.compile("([^\s]*\.(?:png|jpg|jpeg|gif|PNG|JPG|JPEG|GIF))")
slidderpath_regex = re.compile(rf"path={img_file_regex.pattern}")
document_id_regex = re.compile("1[a-zA-Z0-9-_]{43}")


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

    def get_presentation(self, presentation_id_or_name):
        if document_id_regex.match(presentation_id_or_name) is not None:
            presentation_id = presentation_id_or_name
        else:
            response = self.drive_service.files().list(
                q=f"name='{presentation_id_or_name}'",
            ).execute()
            files = [f["id"] for f in response["files"]]
            if len(files) == 0:
                raise RuntimeError(f"Found no presentations matching '{presentation_id_or_name}'")
            if len(files) > 1:
                links = ", ".join(f"https://drive.google.com/open?id={f}" for f in files)
                raise RuntimeError(
                    f"Found multiple files matching '{presentation_id_or_name}': {links}"
                )
            presentation_id = files[0]
        return self.slides_service.presentations().get(presentationId=presentation_id).execute()

    def upload_image(self, localpath, verbose=False):
        file_metadata = {
            "name": localpath,
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
        if verbose:
            print(f"Uploaded {localpath} -> {obj_id}")
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


def get_file_hash(filename, buf_size=65536):
    hasher = hashlib.sha1()
    if filename.startswith("http"):
        fp = urllib.request.urlopen(filename)
    else:
        fp = open(filename, "rb")
    while True:
        data = fp.read(buf_size)
        if not data:
            break
        hasher.update(data)
    fp.close()
    return hasher.hexdigest()


def main(
    presentation_id_or_name,
    directory="./",
    debug=True,
):
    """slidder: the Google Slides image updater

    Given a presentation `presentation_id_or_name` (either a 44-character string starting with 1
    that you can find in the URL of your presentation) or a document name (must be unique),
    replace annotated images with files from the given `directory`.
    Images are annotated with a string pointing to their path with a string like
    "path=./my/path/image.png" in the image description, which you see when you
    right click on the image in Slides -> Alt text -> Description.
    """
    gapi = GAPI()
    presentation = gapi.get_presentation(presentation_id_or_name)
    slides = presentation.get("slides")
    if debug:
        print(f"The presentation contains {len(slides)} slides")

    requests = []
    uploaded_files = {}
    for s, slide in enumerate(slides, start=1):
        images = [el for el in slide.get("pageElements", []) if "image" in el]
        slide_id = slide.get("objectId")
        for image in images:
            img_id = image.get("objectId")
            title = image.get("title", "")
            desc = image.get("description", "")
            url = image.get("image", {}).get("contentUrl")
            annotated_files = slidderpath_regex.findall(desc)
            if len(annotated_files) != 1:
                if len(annotated_files) > 1:
                    print(
                        f"Slide #{s} has an image with multiple definitions of path=...: "
                        f"{'; '.join(annotated_files)} -- skipping!"
                    )
                continue
            raw_fname = annotated_files[0]
            fname = os.path.join(directory, raw_fname)
            if debug:
                print(f"{img_id} on slide #{s} ({slide_id}): {url} -> {raw_fname} -> {fname}")
            if not os.path.isfile(fname):
                print(
                    f"Tried to find {fname} (on slide #{s}), but didn't find it "
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
            requests.append({  # replaceImage resets these to "" for some reason
                "updatePageElementAltText": {
                    "objectId": img_id,
                    "title": title,
                    "description": desc,
                },
            })

    drive_file_ids = [f["file_id"] for f in uploaded_files.values()]
    if drive_file_ids and requests:
        # make all uploaded files world-readable-if-have-link so that we can import them
        gapi.make_public(drive_file_ids)

    response = None
    if requests:
        if debug:
            print(f"Have requests:\n{requests}")
        gapi.slides_service.presentations().batchUpdate(
            presentationId=presentation["presentationId"], body={"requests": requests},
        ).execute()

    gapi.remove_files(drive_file_ids)  # remove all uploaded files from the app directory


def id_images(
    presentation_id_or_name,
    directory=".",
):
    """Find images in the presentation and try to match them with local files.
    """
    raise RuntimeError("Not implemented")

    gapi = GAPI()
    presentation = gapi.get_presentation(presentation_id_or_name)
    slides = presentation.get("slides")

    # find hashes for all local files
    hashes = {}
    for dirpath, dirnames, filenames in os.walk(directory):
        for f in filenames:
            f = os.path.join(dirpath, f)
            if img_file_regex.match(f) is not None:
                file_hash = get_file_hash(f)
                if file_hash in hashes:
                    print(
                        f"WARNING: have two files with hash {file_hash}: "
                        f"{hashes[file_hash]} and {f}; using the first one to annotate"
                    )
                    continue
                hashes[file_hash] = f
    for h, f in sorted(hashes.items()):
        print(f"{h}\t{f}")

    for s, slide in enumerate(slides, start=1):
        images = [el for el in slide.get("pageElements", []) if "image" in el]
        slide_id = slide.get("objectId")
        for image in images:
            img_id = image.get("objectId")
            title = image.get("title", "")
            desc = image.get("description", "")
            url = image.get("image", {}).get("contentUrl")
            annotated_files = slidderpath_regex.findall(desc)
            if len(annotated_files) > 1:
                print(
                    f"WARNING: Slide #{s} has an image with multiple definitions of path=...: "
                    f"{'; '.join(annotated_files)}"
                )
                continue
            img_hash = get_file_hash(url)
            print(f"Found image on slide {s}: hashes to {img_hash} (URL: {url})")
            # hmm, looks like the images get processed a bit, so this hash-based approach
            # won't work
            # TODO

            if len(annotated_files) == 1:
                # already have an annotation: need to check it's the same
                pass


if __name__ == "__main__":
    fire.Fire(main)
