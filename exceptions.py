""" These exceptions may be raised at different stages of the scheduling process """

class ParameterStoreError(Exception):
    """Custom exception raised when there is a problem with Parameter Store"""

class SharepointInitError(Exception):
    """Custom exception raised when Sharepoint fails to initialize"""

class SharepointColumnMappingError(Exception):
    """Custom exception raised when Sharepoint columns could not be fully mapped to the schema"""

class WebexIntegrationInitError(Exception):
    """Custom exception raised when Webex Integration fails to initialize"""

class WebexBotInitError(Exception):
    """Custom exception raised when Webex Bot fails to initialize"""
