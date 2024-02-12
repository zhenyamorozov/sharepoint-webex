"""
    Utilities for interacting with Parameter Store. Implemented for AWS SSM Parameter Store.
"""

import json
import time

import webexteamssdk

import boto3


def getSharepointParams():
    """Returns the saved Sharepoint parameters from Parameter Store, related to the current working Lists folder.

    Args:
        None
    Returns:
        Tuple of three values:
            - spSiteURL: Sharepoint site URL
            - spListName: Sharepoint Lists list name
            - spFolderName: the name of the current working Lists folder
    """

    # load parameters from parameter store
    ssm_client = boto3.client("ssm")
    
    spSiteURL = ssm_client.get_parameter(
        Name="/sharepoint-webex/spSiteURL",
        WithDecryption=True
    )['Parameter']['Value']
    spListName = ssm_client.get_parameter(
        Name="/sharepoint-webex/spListName",
        WithDecryption=True
    )['Parameter']['Value']
    spFolderName = ssm_client.get_parameter(
        Name="/sharepoint-webex/spFolderName",
        WithDecryption=True
    )['Parameter']['Value']

    ssm_client.close()
    return (spSiteURL, spListName, spFolderName)


def saveSharepointParams(spSiteURL, spListName, spFolderName):
    """Saves current working Sharepoint Lists parameters: site URL, list name, folder name to Parameter Store

    Args:
        spSiteURL (str)
        spListName (str)
        spFolderName (str)
    Returns:
        None
    """
    ssm_client = boto3.client("ssm")
    ssm_client.put_parameter(
        Name="/sharepoint-webex/spSiteURL",
        Value=spSiteURL,
        Type="String",
        Overwrite=True
    )
    ssm_client.put_parameter(
        Name="/sharepoint-webex/spListName",
        Value=spListName,
        Type="String",
        Overwrite=True
    )
    ssm_client.put_parameter(
        Name="/sharepoint-webex/spFolderName",
        Value=spFolderName,
        Type="String",
        Overwrite=True
    )


def getWebexIntegrationToken(webex_integration_client_id, webex_integration_client_secret):
    """Returns a fresh, usable Webex Integration access token.

    Webex Integration access tokens are acquired through OAuth and must be refreshed regularly.
    OAuth-provided access token and refresh token have limited lifetimes. As of now,
        access_token lifetime is 14 days since creation
        refresh_token lifetime is 90 days since last use
    This function reads tokens from Parameter Store, refreshes the access_token if it's more than halftime old,
    and returns the access_token.

    Args:
        webex_integration_client_id - used if access token refresh is needed
        webex_integration_client_secret - used if access token refresh is needed

    Returns:
        accessToken: fresh, usable Webex Integration access token
    """

    # read access tokens from Parameter Store
    ssm_client = boto3.client("ssm")
    ssmStoredParameter = ssm_client.get_parameter(
        Name="/sharepoint-webex/webexTokens",
        WithDecryption=True
    )
    currentTokens = json.loads(ssmStoredParameter['Parameter']['Value'])
    accessToken = currentTokens['access_token']
    createdTime = currentTokens['created']
    lifetime = 14*24*60*60    # 14 days
    if createdTime + lifetime/2 < time.time():
        # refresh token
        refreshToken = currentTokens['refresh_token']

        webexApi = webexteamssdk.WebexTeamsAPI(access_token=accessToken)    # passing expired access_token should still work, the API object can be initiated with any string
        newTokens = webexApi.access_tokens.refresh(
            client_id=webex_integration_client_id,
            client_secret=webex_integration_client_secret,
            refresh_token=refreshToken
        )
        # save the new access token to the Parameter Store
        saveWebexIntegrationTokens(dict(newTokens.json_data))

        accessToken = newTokens.access_token

    ssm_client.close()
    return accessToken


def saveWebexIntegrationTokens(tokens):
    """Saves Webex Integration tokens to Parameter Store. Adds `created` timestamp for token lifetime tracking.

    Args:
        tokens: dict of Webex Integration tokens data, as it comes from the API call

    Returns:
        None
    """
    tokens['created'] = time.time()

    ssm_client = boto3.client("ssm")
    ssmStoredParameter = ssm_client.put_parameter(
        Name="/sharepoint-webex/webexTokens",
        Value=json.dumps(tokens),
        Type="SecureString",
        Overwrite=True
    )
    ssm_client.close()
    return ssmStoredParameter
