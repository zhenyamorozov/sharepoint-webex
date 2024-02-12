"""
This implements the Webex Webinar scheduling process.
Can be launched as a standalone script or invoked by the control bot command.
"""

# common imports
import os
import logging
import io
import json
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
from email.utils import getaddresses
import tempfile

# SDK imports
import webexteamssdk
from office365.sharepoint.client_context import ClientContext, ClientCredential
from office365.sharepoint.listitems.caml.query import CamlQuery

# my imports
from param_store import (
    getSharepointParams,
    getWebexIntegrationToken
)
from exceptions import (
    ParameterStoreError,
    SharepointInitError,
    SharepointColumnMappingError,
    WebexIntegrationInitError,
    WebexBotInitError
)


def loadParameters(logger):
    """
        First step in the scheduling process.

        Loads credentials and other parameters from env variables.
        Checks if all required values are provided.
        Changes global variables.

        Args:
            None

        Returns:
            None
    """
    # Required user-set parameters from the env
    global SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_LIST_NAME
    global WEBEX_INTEGRATION_CLIENT_ID, WEBEX_INTEGRATION_CLIENT_SECRET
    global WEBEX_BOT_TOKEN, WEBEX_BOT_ROOM_ID
    # Optional user-set parameters from the env
    global SHAREPOINT_PARAMS, WEBEX_INTEGRATION_PARAMS

    load_dotenv(override=True)

    # load required parameters from env
    logger.info("Loading required parameters from env.")

    SHAREPOINT_CLIENT_ID = os.getenv('SHAREPOINT_CLIENT_ID')
    if not SHAREPOINT_CLIENT_ID:
        logger.fatal("Sharepoint Client ID is missing. Provide with SHAREPOINT_CLIENT_ID environment variable.")
        raise SystemExit()
    SHAREPOINT_CLIENT_SECRET = os.getenv('SHAREPOINT_CLIENT_SECRET')
    if not SHAREPOINT_CLIENT_SECRET:
        logger.fatal("Sharepoint Client Secret is missing. Provide with SHAREPOINT_CLIENT_SECRET environment variable.")
        raise SystemExit()

    WEBEX_INTEGRATION_CLIENT_ID = os.getenv('WEBEX_INTEGRATION_CLIENT_ID')
    if not WEBEX_INTEGRATION_CLIENT_ID:
        logger.fatal("Webex Integration Client ID is missing. Provide with WEBEX_INTEGRATION_CLIENT_ID environment variable.")
        raise SystemExit()
    WEBEX_INTEGRATION_CLIENT_SECRET = os.getenv('WEBEX_INTEGRATION_CLIENT_SECRET')
    if not WEBEX_INTEGRATION_CLIENT_SECRET:
        logger.fatal("Webex Integration Client Secret is missing. Provide with WEBEX_INTEGRATION_CLIENT_SECRET environment variable.")
        raise SystemExit()

    WEBEX_BOT_TOKEN = os.getenv('WEBEX_BOT_TOKEN')
    if not WEBEX_BOT_TOKEN:
        logger.fatal("Webex Bot access token is missing. Provide with WEBEX_BOT_TOKEN environment variable.")
        raise SystemExit()
    WEBEX_BOT_ROOM_ID = os.getenv('WEBEX_BOT_ROOM_ID')
    if not WEBEX_BOT_ROOM_ID:
        logger.fatal("Webex Bot room ID is missing. It is required for logging and control. Provide with WEBEX_BOT_ROOM_ID environment variable.")
        raise SystemExit()
    
    logger.info("Required parameters are loaded from env.")

    # load optional parameters from env
    SHAREPOINT_PARAMS = {}
    SHAREPOINT_PARAMS['columns'] = {
        'create': "Create",
        'startdatetime': "Start Date and Time",
        'duration': "Duration",
        'title': "Title",
        'agenda': "Agenda",
        'cohosts': "Cohosts",
        'panelists': "Panelists",
        'webinarId': "Webinar ID",
        'attendeeUrl': "Attendee URL",
        'hostKey': "Host Key",
        'registrantCount': "Registrant Count"

    }
    SHAREPOINT_PARAMS['nicknames'] = {}
    if os.getenv("SHAREPOINT_PARAMS"):
        logger.info("Loading optional Sharepoint parameters from env.")
        try:
            columns = json.loads(os.getenv('SHAREPOINT_PARAMS'))['columns']
            for i in columns:
                SHAREPOINT_PARAMS['columns'][i] = columns[i]
            logger.info("Optional Sharepoint column parameters are loaded from env.")
        except Exception:
            logger.info("Could not load optional Sharepoint column parameters.")
        try:
            SHAREPOINT_PARAMS['nicknames'] = json.loads(os.getenv('SHAREPOINT_PARAMS'))['nicknames']
            logger.info("Optional Sharepoint nicknames parameters are loaded from env.")
        except Exception:
            logger.info("Could not load optional Sharepoint nicknames parameters.")

    else:
        logger.info("No optional Sharepoint parameters set in env.")

    WEBEX_INTEGRATION_PARAMS = {}
    if os.getenv('WEBEX_INTEGRATION_PARAMS'):
        logger.info("Loading optional Webex Integration parameters from env.")
        try:
            WEBEX_INTEGRATION_PARAMS = json.loads(os.getenv('WEBEX_INTEGRATION_PARAMS'))
            logger.info("Optional Webex Integration parameters are loaded from env.")
        except Exception as ex:
            logger.info("Could not load optional Webex Integration parameters. " + str(ex))
    else:
        logger.info("No optional Webex Integration parameters set in env.")


def initSharepoint():
    """Initializes access to Sharepoint.

        Args:
            None

        Returns:
            A tuple of three items (ssApi, spFolder, spColumnMap)
                spList: Sharepoint Lists list object
                spFolder: Sharepoint Lists folder object
                spColumnMap: dict mapping column names to Sharepoint internal field names

        Raises:
            ParameterStoreError: Could not read Sharepoint parameters from parameter store.
            SharepointInitError: An error occurred accessing Sharepoint, or Sharepoint List name or Sharepoint folder name is wrong.
            SharepointColumnMappingError: Sharepoint List columns could not be properly mapped.
             Occurs when one or more of the required columns are missing from the provided Sharepoint List.
    """
    global spColumnMap

    try:
        spSiteURL, spListName, spFolderName = getSharepointParams()
    except Exception as ex:
        raise ParameterStoreError("Could not read Sharepoint parameters from parameter store. " + str(ex))

    try:
        spApi = ClientContext(spSiteURL).with_credentials(
            ClientCredential(
                SHAREPOINT_CLIENT_ID, 
                SHAREPOINT_CLIENT_SECRET
            )
        )

        # query the list
        spList = spApi.web.lists.get_by_title(spListName).get().execute_query()

        # fetch folder and find the working folder
        qry = CamlQuery()
        qry.ViewXml = """
        <View Scope='DefaultValue'>
            <Query>
                <Where>
                    <BeginsWith>
                    <FieldRef Name='ContentTypeId' />
                    <Value Type='ContentTypeId'>0x0120</Value>
                    </BeginsWith>
                </Where>
            </Query>
        </View>    
        """
        folders = spList.get_items(qry).execute_query()
        spFolder = next(f for f in folders if f.properties['Title']==spFolderName)

        # fetch Sharepoint list fields
        fields = spList.fields.get().execute_query()

    except Exception as ex:
        raise SharepointInitError("Could not initialize access to Sharepoint. Check if Sharepoint parameters (Site URL, List Name, Folder Name) are correct." + str(ex))

    columns = {}
    # Build column map for later reference - translates column names to Sharepoint column internal names
    for column in fields:
        columns[column.title] = column.internal_name  #column.id

    spColumnMap = {}
    for column in SHAREPOINT_PARAMS['columns']:
        if SHAREPOINT_PARAMS['columns'][column] in columns:
            spColumnMap[column] = columns[SHAREPOINT_PARAMS['columns'][column]]
    # check if all required columns are present in the Sharepoint list
    requiredColumns = ['create', 'startdatetime', 'title', 'webinarId']
    columnsDiff = set(requiredColumns) - set(spColumnMap.keys())
    if columnsDiff:
        raise SharepointColumnMappingError("Some required column(s) is(are) missing in your Sharepoint list: " + ", ".join(columnsDiff))

    return (spList, spFolder, spColumnMap)


def initWebexIntegration():
    """Initializes access to Webex integration.

        Args:
            None

        Returns:
            webexApi: webexteamssdk.WebexTeamsAPI object to control access to Webex Integration

        Raises:
            WebexIntegrationInitError
    """
    try:
        # get a fresh token
        webexToken = getWebexIntegrationToken(WEBEX_INTEGRATION_CLIENT_ID, WEBEX_INTEGRATION_CLIENT_SECRET)
        # print(webexToken) #TODO dev
        # init the object
        webexApi = webexteamssdk.WebexTeamsAPI(webexToken)
        # check if API is functional by requesting `me`
        assert webexApi.people.me()
    except Exception as ex:
        raise WebexIntegrationInitError(ex)

    return webexApi


def initWebexBot():
    """Initializes access to Webex bot for logging and control.

        Args:
            None

        Returns:
            botApi: webexteamssdk.WebexTeamsAPI object to control access to Webex Integration
    """
    try:
        botApi = webexteamssdk.WebexTeamsAPI(WEBEX_BOT_TOKEN)
        # check if it's a bot
        assert botApi.people.me().type == "bot"
        # check if the bot can access the logging room
        assert botApi.rooms.get(WEBEX_BOT_ROOM_ID)
    except Exception:
        raise WebexBotInitError()

    return botApi


def getWebinarProperty(propertyName, spRow=None):
    """Returns value for webinar property taken from a source, in order of priority:
        1. If propertyName column exists in Sharepoint list and cell is not empty,
            return its value
        2. If propertyName filed exists in WEBEX_INTEGRATION_PARAMS and is not empty,
            return its value
        3. Return None

        Args:
            propertyName: Name of the Webinar property. Naming aligns with API specs
            spRow: Sharepoint row

        Returns:
            propertyValue (str): for number/text fields,
            or None
    """
    
    try:
        # if property value is specified in Sharepoint list, return it
        return spRow.properties[spColumnMap[propertyName]]
    except Exception:
        pass
    try:
        # if property value is specified in env WEBEX_INTEGRATION_PARAMS, return it
        return WEBEX_INTEGRATION_PARAMS[propertyName]
    except Exception:
        # otherwise, return None
        return None


def stringContactsToDict(contacts):
    """Returns dict of contacts from string of contacts

        Args: contacts(str): Comma-separated list of contacts, each contact represented as "name <email>".
            If name is not specified, 'Panelist' is used.
            If email is not specified, the function will try to match contact by nickname configured in env.

        Returns:
            _res: dict of contacts {email: name}

    """
    _res = {}
    for contact in getaddresses(str(contacts).split(",")):
        if "@" in contact[1]:
            # email specified
            _res[contact[1].strip().lower()] = contact[0].strip() or "Panelist"
        else:
            # email not specified
            try:
                _res[SHAREPOINT_PARAMS['nicknames'][contact[1].strip().lower()]['email']] = SHAREPOINT_PARAMS['nicknames'][contact[1].strip().lower()]['name']
            except Exception:
                pass

    return _res


def run():
    """This is the main function that runs the scheduling process. Takes no arguments, returns nothing, just 
    does the scheduling job. It generates a detailed and a brief log and sends them to the Webex space.

        Args: None

        Returns: None

    """    

    #
    #   Initialize logging
    #
    logger = logging.getLogger(__name__)
    # Logger usage:
    # logger.fatal("Message in case of a fatal error causing SystemExit")
    # logger.error("Message in case of an error, goes to the brief log")
    # logger.warning("Message to be output to the brief log (the Webex message itself)")
    # logger.info("Message to be output in the full log (text file attached to Webex message)")
    # logger.debug("Message to be output in the console only")

    # set log level to DEBUG
    logger.setLevel(logging.DEBUG)

    # log handler for brief log
    briefLogString = io.StringIO()
    briefLogHandler = logging.StreamHandler(briefLogString)
    briefLogHandler.setLevel(logging.WARNING)
    logger.addHandler(briefLogHandler)

    # log handler for full log
    fullLogString = io.StringIO()
    fullLogHandler = logging.StreamHandler(fullLogString)
    fullLogHandler.setLevel(logging.INFO)
    logger.addHandler(fullLogHandler)

    # log handler for console
    consoleLogHandler = logging.StreamHandler()
    consoleLogHandler.setLevel(logging.DEBUG)
    logger.addHandler(consoleLogHandler)

    startTime = datetime.now()
    logger.warning("Starting...")

    #
    # Load env variables and check if all env variables are provided
    #
    logger.info("Loading parameters and checking if all required parameters are provided.")
    loadParameters(logger=logger)
    logger.info("Required parameters are successfully loaded.\n")

    #
    # Initialize access to Sharepoint
    #
    logger.info("Initializing access to Sharepoint.")
    try:
        spList, spFolder, spColumnMap = initSharepoint()
    except ParameterStoreError as ex:
        logger.fatal("Could not read Sharepoint Folder Name from Parameter Store. Check local AWS configuration. Service reported: " + str(ex))
        raise SystemExit()
    except SharepointInitError as ex:
        logger.fatal("Sharepoint API connection error. " + str(ex))
        raise SystemExit()
    except SharepointColumnMappingError as ex:
        logger.fatal("Sharepoint List column mapping error. " + str(ex))
        raise SystemExit()
    except Exception as ex:
        logger.fatal("Sharepoint initialization error. " + str(ex))
        raise SystemExit()
    logger.info("Successfully initialized access to Sharepoint.")

    #
    # Initialize access to Webex Integration
    #
    logger.info("Initializing access to Webex Integration.")
    try:
        webexApi = initWebexIntegration()
    except ParameterStoreError as ex:
        logger.fatal("Could not read Webex Integration tokens from Parameter Store. Check local AWS configuration. Service reported: " + str(ex))
        raise SystemExit()
    except WebexIntegrationInitError as ex:
        logger.fatal("Could not initialize Webex Integration. Service reported: " + str(ex))
        raise SystemExit()
    except Exception as ex:
        logger.fatal("Could not initialize Webex Integration. Service reported: " + str(ex))
        raise SystemExit()
    logger.info("Successfully initialized access to Webex Integration.")

    #
    # Initialize access to Webex bot for logging and control
    #
    logger.info("Initializing access to Webex bot.")
    try:
        botApi = initWebexBot()
    except Exception as ex:
        logger.fatal("Could not initialize Webex bot. Service reported: " + str(ex))
        raise SystemExit()
    logger.info("Successfully initialized access to Webex bot.")


    # Set default time zone
    os.environ['TZ'] = 'UTC'

    # Will calculate total amount of registrants
    totalRegistrantCount = 0


    #
    # Loop over the Sharepoint list
    #
    qry = CamlQuery()
    qry.FolderServerRelativeUrl = f"Lists/{spList.title}/{spFolder.properties['Title']}"

    for spRow in spList.get_items(qry).execute_query():

        if spRow.properties[spColumnMap['create']]:
            event = {}

            logger.info("")    # insert empty line into log

            # gather all webinar properties
            event['title'] = getWebinarProperty('title', spRow) or "Generic Webinar Title"
            try:
                event['agenda'] = getWebinarProperty('agenda', spRow)
                event['scheduledType'] = getWebinarProperty('scheduledType', spRow) or 'webinar'
                event['startdatetime'] = datetime.fromisoformat(getWebinarProperty('startdatetime', spRow))
                event['duration'] = getWebinarProperty('duration', spRow) or 60    # by default, set duration to 1 hour
                event['duration'] = int(float(event['duration']))    # make sure it's integer
                event['enddatetime'] = event["startdatetime"] + timedelta(minutes=event['duration'])
                event['timezone'] = getWebinarProperty('timezone', spRow) or "UTC"
                event['siteUrl'] = getWebinarProperty('siteUrl', spRow)    # if not set, Webex default will be used
                event['password'] = getWebinarProperty('password', spRow)    # by default, randomly generated by Webex
                event['panelistPassword'] = getWebinarProperty('panelistPassword', spRow)    # by default, randomly generated by Webex
                event['reminderTime'] = getWebinarProperty('reminderTime', spRow) or 30    # by default, set reminder to go 30 minutes before the session
                # if it's too late to send reminder, skip it
                if datetime.utcnow().replace(tzinfo=timezone.utc) >= event['startdatetime'] - timedelta(minutes=event['reminderTime']):
                    event['reminderTime'] = 0
                event['registration'] = getWebinarProperty('registration', spRow) or \
                    {
                        'autoAcceptRequest': True,
                        'requireFirstName': True,
                        'requireLastName': True,
                        'requireEmail': True
                    }    # registration is enabled by default
                event['enabledJoinBeforeHost'] = getWebinarProperty('enabledJoinBeforeHost', spRow)   # let attendees join before host
                event['joinBeforeHostMinutes'] = getWebinarProperty('joinBeforeHostMinutes', spRow)   # set webinar to start minutes before the scheduled start time

                # add invited cohosts
                event['cohosts'] = getWebinarProperty('cohosts', spRow)
                if not isinstance(event['cohosts'], dict):
                    event['cohosts'] = stringContactsToDict(event['cohosts'])
                # add invited panelists
                event['panelists'] = getWebinarProperty('panelists', spRow)
                if not isinstance(event['panelists'], dict):
                    event['panelists'] = stringContactsToDict(event['panelists'])
                # add panelists which are always invited
                alwaysInvitePanelists = getWebinarProperty('alwaysInvitePanelists')
                alwaysInvitePanelists = stringContactsToDict(alwaysInvitePanelists)
                event['panelists'].update(alwaysInvitePanelists)

                event['id'] = getWebinarProperty('webinarId', spRow)
                logger.info("Processing \"{}\"".format(event['title']))
            except Exception as ex:
                logger.error("Failed to process \"{}\". The webinar property is not valid: {}".format(event['title'], ex))
                continue

            if not event['id']:
                # create event
                try:
                    w = webexApi.meetings.create(
                        title=event['title'],
                        agenda=event['agenda'],
                        scheduledType=event['scheduledType'],
                        start=str(event["startdatetime"]),
                        end=str(event["enddatetime"]),
                        timezone=event['timezone'],
                        siteUrl=event['siteUrl'],
                        password=event['password'],
                        panelistPassword=event['panelistPassword'],
                        reminderTime=event['reminderTime'],
                        registration=event['registration'],
                        enabledJoinBeforeHost=event['enabledJoinBeforeHost'],
                        joinBeforeHostMinutes=event['joinBeforeHostMinutes']
                    )
                    logger.warning("Created webinar {}".format(w.title))
                except Exception as ex:
                    logger.error("Failed to create webinar \"{}\". API returned error: {}".format(event['title'], ex))
                    try:
                        for err in ex.details['errors']:
                            logger.error("  " + err['description'])
                    except Exception:
                        pass
                    continue
                pass

                # update newly created webinar ID and info back into Sharepoint list
                try:
                    if 'webinarId' in spColumnMap:
                        spRow.set_property(spColumnMap['webinarId'], w.id)
                    else:
                        logger.error("No column in Sharepoint list to save Webinar ID.")    # critical for app logic

                    if 'hostKey' in spColumnMap:
                        spRow.set_property(spColumnMap['hostKey'], w.hostKey)
                    else:
                        logger.info("No column in Sharepoint list to save Host Key.")

                    if 'attendeeUrl' in spColumnMap:
                        spRow.set_property(spColumnMap['attendeeUrl'], "Manually copy the Attendee URL from Webex")
                        # TODO to be added when API supports - Attendee Registration URL
                    else:
                        logger.info("No column in Sharepoint list to save Attendee URL.")

                    spRow.update().execute_query()
                    logger.info("Updated webinar information into Sharepoint list.")

                except Exception as ex:
                    logger.error("Failed to update created webinar information into Sharepoint list. API returned error: {}".format(ex))

            else:
                # update existing event
                try:
                    w = webexApi.meetings.get(event['id'])

                    needUpdateSendEmail = \
                        event['title'] != w.title \
                        or event['startdatetime'] != datetime.fromisoformat(w.start) \
                        or event['enddatetime'] != datetime.fromisoformat(w.end)    # fromisoformat() cannot process ISO-8601 strings prior to Python 3.11, thus remove the 'Z'

                    needUpdate = \
                        needUpdateSendEmail \
                        or event['agenda'] != w.agenda

                    if needUpdate:
                        w = webexApi.meetings.update(
                            meetingId=event['id'],
                            title=event['title'],
                            agenda=event['agenda'],
                            scheduledType=event['scheduledType'],
                            start=str(event["startdatetime"]),
                            end=str(event["enddatetime"]),
                            # timezone=event['timezone'],
                            # siteUrl=event['siteUrl'],
                            password=event['password'] or w.password,    # password is required for update()
                            panelistPassword=event['panelistPassword'],
                            # reminderTime=event['reminderTime'],
                            # registration=event['registration'],
                            enabledJoinBeforeHost=event['enabledJoinBeforeHost'],
                            joinBeforeHostMinutes=event['joinBeforeHostMinutes'],                            
                            sendEmail=needUpdateSendEmail
                        )
                        logger.warning("Updated webinar information: {}".format(w.title))
                except Exception as ex:
                    logger.error("Failed to update webinar \"{}\". API returned error: {}".format(event['title'], ex))
                    try:
                        for err in ex.details['errors']:
                            logger.error("  " + err['description'])
                    except Exception:
                        pass
                    continue

                # refresh webinar registrant count in Sharepoint list
                try:
                    registrantCount = sum(1 for _ in webexApi.meeting_invitees.list(w.id))
                    # TODO implement in a more efficient way once the list-meeting-registrants endpoint is added to the SDK
                    totalRegistrantCount += registrantCount

                    if 'registrantCount' in spColumnMap:
                        spRow.set_property(spColumnMap['registrantCount'], registrantCount)
                    else:
                        raise Exception("No column in Sharepoint list to save Registration Count.")

                    spRow.update().execute_query()
                    logger.info("Refreshed webinar Registration Count in Sharepoint list.")

                except Exception as ex:
                    logger.error("Failed to refresh webinar Registration Count in Sharepoint list. API returned error: {}".format(ex))

            # update invitees (panelists and cohosts) for created or updated event
            try:
                # collect currently invited panelists and cohosts
                # also serves as an "uninvite list" - checked invitees are removed from the list
                # if there are any remaining, they will be uninvited
                currentInvitees = {}
                for i in webexApi.meeting_invitees.list(w.id, panelist=True):
                    if i.panelist or i.coHost:
                        currentInvitees[i.email] = i
            except Exception as ex:
                logger.error("Failed to process invitees for webinar \"{}\". API returned error: {}".format(event['title'], ex))
            else:

                # process panelists and cohosts

                if getWebinarProperty('noCohosts'):
                    # treat cohosts as panelists
                    event['panelists'].update(event['cohosts'])
                    event['cohosts'] = {}

                eventInvitees = event['panelists'] | event['cohosts']    # merged dicts: https://peps.python.org/pep-0584/
                for email in eventInvitees:
                    if email in currentInvitees:
                        # already invited
                        if eventInvitees[email] != currentInvitees[email].displayName \
                                or (email in event['cohosts']) != currentInvitees[email].coHost:
                            # name or status changed
                            try:
                                r = webexApi.meeting_invitees.update(
                                    meetingInviteeId=currentInvitees[email].id,
                                    email=email,
                                    displayName=eventInvitees[email],
                                    panelist=email in event['panelists'] or email in event['cohosts'],    # cohosts must also be panelists as per Webex API behavior
                                    coHost=email in event['cohosts'],
                                    sendEmail=True
                                )
                                logger.info("Updated invitee {} <{}>".format(eventInvitees[email], email))
                            except Exception as ex:
                                logger.error("Failed to update invitee \"{}\" for webinar \"{}\". API returned error: {}".format(email, event['title'], ex))
                        del currentInvitees[email]    # remove processed from the uninvite list
                    else:
                        # new, need to invite
                        try:
                            r = webexApi.meeting_invitees.create(
                                meetingId=w.id,
                                email=email,
                                displayName=eventInvitees[email],
                                panelist=email in event['panelists'] or email in event['cohosts'],    # cohosts must also be panelists as per Webex API behavior
                                coHost=email in event['cohosts'],
                                sendEmail=True
                            )
                            logger.info("Invited {} <{}>".format(eventInvitees[email], email))
                        except Exception as ex:
                            logger.error("Failed to create invitee \'{}\' for webinar \"{}\". API returned error: {}".format(email, event['title'], ex))
                # uninvite panelists/cohosts who remained in the uninvite list
                for email in currentInvitees:
                    try:
                        r = webexApi.meeting_invitees.delete(
                            meetingInviteeId=currentInvitees[email].id
                        )
                        logger.info("Uninvited {} <{}>".format(currentInvitees[email].displayName, email))
                    except Exception as ex:
                        logger.error("Failed to delete invitee \"{}\" from webinar \"{}\". API returned error: {}".format(email, event['title'], ex))

    # /for
    
    logger.warning(f"Done in {datetime.now()-startTime}. Total registrants: {totalRegistrantCount}.")

    #
    # Process logs and close logging
    #
    try:
        with tempfile.NamedTemporaryFile(
            prefix=datetime.utcnow().strftime("%Y%m%d-%H%M%S "),
            suffix=".txt",
            mode="wt",
            encoding="utf-8",
            delete=False
        ) as tmp:
            tmp.write(fullLogString.getvalue())

        botApi.messages.create(
            roomId=WEBEX_BOT_ROOM_ID,
            text="Done creating and updating webinars. Full log attached. Brief log follows.\n\n" + briefLogString.getvalue(),
            files=[tmp.name]
        )

        os.remove(tmp.name)
    except Exception as ex:
        logger.error("Failed to post log into Webex bot room. " + str(ex))

    # close logging
    briefLogString.close()
    fullLogString.close()
    logger.removeHandler(briefLogHandler)
    logger.removeHandler(fullLogHandler)
    logger.removeHandler(consoleLogHandler)
    logging.shutdown()


# Run the scheduling process if launched as a script
if __name__ == "__main__":
    run()
