# 📊 Google Analytics 4 Data Exporter

## Technologies Used

* **Language**: C#
* **IDE**: Visual Studio 2022
* **Framework**: .NET 6
* **API**: Google Analytics Data API v1

## Prerequisites

* Access to a **Google Analytics 4 property**
* A **`client_secrets.json` service account file** (generated from Google Cloud)
* **GA4 API enabled** in your Google Cloud project

---

## How to Set Up a Connection to the Google Analytics 4 API

### 1. Install NuGet Packages

```bash
dotnet add package Google.Analytics.Data.V1Beta --version 2.0.0-beta09
dotnet add package ClosedXML --version 0.105.0
```

### 2. Create a Google Service Account

1. Go to [console.cloud.google.com](https://console.cloud.google.com/)
2. Select or create a **Google Cloud project**
3. In the **"IAM"** section, click **"Service Accounts" > "+ Create Service Account"**
4. Once created, an email address is generated (keep it for later), then click **"Manage Keys"** (the three dots in the "Actions" tab)
5. Create a key of type `.json`. The file will be downloaded automatically.
6. Finally, enable the **Google Analytics Data API**:
   -> [Enable here](https://console.developers.google.com/apis/api/analyticsdata.googleapis.com/overview)

### 3. Add the Service Account to GA4

1. Go to [analytics.google.com](https://analytics.google.com/)
2. Access the **GA4 property**:
   Click **Admin > Property > Property Access Management**
3. Click **Add User**, then paste the service account email address
   (found in your JSON file, e.g., `service-account@XXXX.iam.gserviceaccount.com`)
4. Assign the **"Viewer"** role *(minimum)*

### 4. Generate a JSON Key and Add It as an Environment Variable

Once the API is enabled and the service account is authorized, you must **set the path** to the JSON file as an environment variable.

On **Windows**:

* Go to **Settings > Edit environment variables for your account**
* Click **New user variable**
* Enter:

  * **Name**: `GOOGLE_APPLICATION_CREDENTIALS` *(required)*
  * **Value**: path to the `.json` file

```
Example:
Name: GOOGLE_APPLICATION_CREDENTIALS
Value: C:\Users\DevIt\Desktop\GoogleAnalytics\GoogleAnalytics\bin\Debug\net6.0\fileGA4-.json
```
