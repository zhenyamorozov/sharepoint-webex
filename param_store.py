""" Utilities for interacting with Parameter Store. Supports AWS SSM Parameter Store and AWS Secrets Manager. """

import json
import os
import time

import webexteamssdk
import boto3


def _use_secrets_manager():
    """Check if Secrets Manager should be used based on environment variable."""
    return os.getenv('AWS_SECRET_STORE', 'ssm').lower() == 'secretsmanager'


def _get_secret_value(name):
    """Get value from Secrets Manager."""
    client = boto3.client('secretsmanager')
    response = client.get_secret_value(SecretId=name)
    client.close()
    return response['SecretString']


def _put_secret_value(name, value):
    """Put value to Secrets Manager."""
    client = boto3.client('secretsmanager')
    try:
        client.update_secret(SecretId=name, SecretString=value)
    except client.exceptions.ResourceNotFoundException:
        client.create_secret(Name=name, SecretString=value)
    client.close()


def _get_parameter_value(name):
    """Get value from SSM Parameter Store."""
    client = boto3.client('ssm')
    response = client.get_parameter(Name=name, WithDecryption=True)
    client.close()
    return response['Parameter']['Value']


def _put_parameter_value(name, value, secure=False):
    """Put value to SSM Parameter Store."""
    client = boto3.client('ssm')
    param_type = 'SecureString' if secure else 'String'
    client.put_parameter(Name=name, Value=value, Type=param_type, Overwrite=True)
    client.close()


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
    if _use_secrets_manager():
        spSiteURL = _get_secret_value('/sharepoint-webex/spSiteURL')
        spListName = _get_secret_value('/sharepoint-webex/spListName')
        spFolderName = _get_secret_value('/sharepoint-webex/spFolderName')
    else:
        spSiteURL = _get_parameter_value('/sharepoint-webex/spSiteURL')
        spListName = _get_parameter_value('/sharepoint-webex/spListName')
        spFolderName = _get_parameter_value('/sharepoint-webex/spFolderName')
    
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
    if _use_secrets_manager():
        _put_secret_value('/sharepoint-webex/spSiteURL', spSiteURL)
        _put_secret_value('/sharepoint-webex/spListName', spListName)
        _put_secret_value('/sharepoint-webex/spFolderName', spFolderName)
    else:
        _put_parameter_value('/sharepoint-webex/spSiteURL', spSiteURL)
        _put_parameter_value('/sharepoint-webex/spListName', spListName)
        _put_parameter_value('/sharepoint-webex/spFolderName', spFolderName)


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

    if _use_secrets_manager():
        token_data = _get_secret_value('/sharepoint-webex/webexTokens')
    else:
        token_data = _get_parameter_value('/sharepoint-webex/webexTokens')
    
    currentTokens = json.loads(token_data)
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

    return accessToken


def saveWebexIntegrationTokens(tokens):
    """Saves Webex Integration tokens to Parameter Store. Adds `created` timestamp for token lifetime tracking.

    Args:
        tokens: dict of Webex Integration tokens data, as it comes from the API call

    Returns:
        None
    """
    tokens['created'] = time.time()
    token_json = json.dumps(tokens)
    
    if _use_secrets_manager():
        _put_secret_value('/sharepoint-webex/webexTokens', token_json)
    else:
        _put_parameter_value('/sharepoint-webex/webexTokens', token_json, secure=True)
