""" This module implements Webex bot functionality. Imported from web.py """

import os

from flask import request, url_for

import webexteamssdk

from webexteamssdk import WebexTeamsAPI
from webexteamssdk.models.cards.card import AdaptiveCard
from webexteamssdk.models.cards.inputs import * # pylint: disable=wildcard-import,unused-wildcard-import
from webexteamssdk.models.cards.components import * # pylint: disable=wildcard-import,unused-wildcard-import
from webexteamssdk.models.cards.container import * # pylint: disable=wildcard-import,unused-wildcard-import
from webexteamssdk.models.cards.actions import * # pylint: disable=wildcard-import,unused-wildcard-import
from webexteamssdk.models.cards.options import * # pylint: disable=wildcard-import,unused-wildcard-import

import schedule

from param_store import (
    getSharepointParams,
    saveSharepointParams,
    getWebexIntegrationToken
)


def init(webAppPublicUrl):
    """ initializes this module, called immediately after importing """
    global botApi

    # initialize Webex Teams bot control object
    try:
        botApi = WebexTeamsAPI(os.getenv("WEBEX_BOT_TOKEN"), wait_on_rate_limit=False)    # this will not raise an exception, even if bot token isn't correct, so,
        # need to make an API call to check if API is functional
        assert botApi.people.me()
    except Exception:
        print("Could not initialize Webex bot API object.")
        raise SystemExit()

    webhookTargetUrl = webAppPublicUrl + "/webhook"

    # delete ALL "Sharepoint-Webex bot..." current webhooks - this bot is supposed to be used only with one instance of the app
    try:
        for wh in botApi.webhooks.list():
            if wh.name.startswith("Sharepoint-Webex bot"):
                botApi.webhooks.delete(wh.id)
    except Exception:
        print("Could not clean up Webex bot API webhooks.")
        raise SystemExit()

    # create new webhooks
    try:
        botApi.webhooks.create(
            name="Sharepoint-Webex bot - attachmentActions",
            targetUrl=webhookTargetUrl,
            resource="attachmentActions",
            event="created",
            filter="roomId=" + os.getenv("WEBEX_BOT_ROOM_ID")
        )
    except Exception:
        print("Could not create a Webex bot API webhook.")
    try:
        botApi.webhooks.create(
            name="Sharepoint-Webex bot - messages",
            targetUrl=webhookTargetUrl,
            resource="messages",
            event="created",
            filter="roomId=" + os.getenv("WEBEX_BOT_ROOM_ID")
        )
    except Exception:
        print("Could not create a Webex bot API webhook.")


# @application.route("/webhook", methods=['GET', 'POST'])
def webhook():
    """ Handles all incoming Webex webhooks """
    # print ("Webhook arrived.")
    # print(request)

    webhookJson = request.json

    # check if the received webhook is properly formed and relevant
    try:
        assert webhookJson['resource'] in ("messages", "attachmentActions")
        assert webhookJson['event'] == "created"
        assert webhookJson['data']['roomId'] == os.getenv("WEBEX_BOT_ROOM_ID")
    except Exception:
        print("The arrived webhook is malformed or does not indicate an actionable event in the log and control room")
        return "Webhook processed."

    # will need our own name
    me = botApi.people.me()

    # static adaptive card - greeting
    greetingCard = AdaptiveCard(
        fallbackText=f"Hi, I am {me.nickName}, I automatically create Webex Webinar sessions based on information in a Sharepoint Lists folder. Adaptive cards feature is required to use me.",
        body=[
            TextBlock(
                text="Sharepoint to Webex Webinar automation",
                weight=FontWeight.BOLDER,
                size=FontSize.MEDIUM,
            ),
            TextBlock(
                text=f"Hi, I am {me.nickName}, I automatically create Webex Webinar sessions based on information in a Sharepoint Lists folder.",
                wrap=True,
            )

        ],
        actions=[
            Submit(title="Schedule now", data={'act': "schedule now"}),
            Submit(title="Set Sharepoint", data={'act': "set sharepoint"}),
            Submit(title="Authorize Webex", data={'act': "authorize webex"}),
            Submit(title="?", data={'act': "help"}),
        ]
    )

    # if webhook indicates a message sent by us (the bot itself), ignore it
    if webhookJson['data']['personId'] == me.id:
        return "Webhook processed."

    # received a text message
    if webhookJson['resource'] == "messages":
        # retrieve the new message details
        # message = botApi.messages.get(webhookJson['data']['id'])
        # print(message)

        # respond with the greeting card to any message
        botApi.messages.create(
            text=greetingCard.fallbackText,
            roomId=os.getenv("WEBEX_BOT_ROOM_ID"),
            attachments=[greetingCard]
        )

    # received a card action
    elif webhookJson['resource'] == "attachmentActions":

        # retrieve the new attachment action details
        action = botApi.attachment_actions.get(webhookJson['data']['id'])
        # print("Action:\n", action)

        # print("actionInputs", actionInputs)

        # "?" (help) action
        if action.type == "submit" and action.inputs['act'] == "help":
            botApi.messages.create(
                markdown="""
Sharepoint and Webex Automation creates webinars in Webex Webinar based on information in Sharepoint Lists.
It is easy to use:
1. Collaborate with your team on webinar planning in a Sharepoint List. Use one folder per webinar series, one list item per webinar. When ready for creation, check **Create**.
2. Click **Schedule Now** button to start webinar scheduling process.
3. Webinars are created.

Features and basic usage: https://github.com/zhenyamorozov/sharepoint-webex#tbd
How to set up and get started: https://github.com/zhenyamorozov/sharepoint-webex/blob/master/docs/get_started.rst#tbd
            """,
                roomId=os.getenv("WEBEX_BOT_ROOM_ID")
            )
            # resend greeting card
            botApi.messages.create(
                text=greetingCard.fallbackText,
                roomId=os.getenv("WEBEX_BOT_ROOM_ID"),
                attachments=[greetingCard]
            )

        # "Schedule now" action
        if action.type == "submit" and action.inputs['act'] == "schedule now":
            try:
                actor = botApi.people.get(personId=webhookJson['actorId'])
                botApi.messages.create(
                    markdown=f"Webinar scheduling requested by <@personId:{actor.id}|{actor.firstName}>. Will start the process. It will take a minute.",
                    roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                )
            except Exception:
                botApi.messages.create(
                    markdown="Webinar scheduling requested. Will start the process. It will take a minute.",
                    roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                )

            # invoke the webinar scheduling process
            schedule.run()

            # send reduced greeting card - only action buttons
            botApi.messages.create(
                text=greetingCard.fallbackText,
                roomId=os.getenv("WEBEX_BOT_ROOM_ID"),
                attachments=[AdaptiveCard(fallbackText=greetingCard.fallbackText, actions=greetingCard.actions)]
            )

        # "Set Sharepoint" action
        if action.type == "submit" and action.inputs['act'] == "set sharepoint":
            try:
                # load current Sharepoint parameters from parameter store
                spSiteURL, spListName, spFolderName = getSharepointParams()

            except Exception:
                spSiteURL = spListName = spFolderName = ""

            card = AdaptiveCard(
                fallbackText="Adaptive cards feature is required to use me.",
                body=[
                    TextBlock(
                        text="Sharepoint setting",
                        weight=FontWeight.BOLDER,
                        size=FontSize.MEDIUM,
                    ),
                    TextBlock(
                        text="Information for scheduled webinars is taken from a Sharepoint Lists folder. Here you can change the Sharepoint Site URL, the name of the List, and the name of the list folder.",
                        wrap=True,
                    ),
                    FactSet(
                        facts=[
                            Fact(
                                title="Sharepoint site URL",
                                value=spSiteURL if spSiteURL else "_none_"
                            ),
                            Fact(
                                title="List name",
                                value=spListName if spListName else "_none_"
                            ),
                            Fact(
                                title="Folder name",
                                value=spFolderName if spFolderName else "_none_"
                            ),
                        ]
                    )
                ],
                actions=[
                    ShowCard(
                        title="Change",
                        card=AdaptiveCard(
                            body=[
                                TextBlock(
                                    text="Change the Sharepoint Lists parameters here. The list should belong to the site. The folder must be located in the root of the list. Nested folders are not supported.",
                                    wrap=True,
                                ),
                                Text('spSiteURL', value=spSiteURL, placeholder="New Sharepoint site URL", isMultiline=False),
                                Text('spListName', value=spListName, placeholder="New Sharepoint List name", isMultiline=False),
                                Text('spFolderName', value=spFolderName, placeholder="New Sharepoint Lists folder name", isMultiline=False),
                            ],
                            actions=[
                                Submit(title="OK", data={'act': "save sharepoint"}),
                            ]
                        )
                    ),
                    # no need for template support
                    # Submit(title="Create Template", data={'act': "create sharepoint template"})
                ]
            )

            # print(card.to_json())
            botApi.messages.create(
                text="Could not send the action card",
                roomId=os.getenv("WEBEX_BOT_ROOM_ID"),
                attachments=[card]
            )

        # "Save Sharepoint" action
        if action.type == "submit" and action.inputs['act'] == "save sharepoint":
            # print(action)

            if 'spSiteURL' not in action.inputs or not action.inputs['spSiteURL'].strip():
                botApi.messages.create(
                    text="Sharepoint site URL cannot be empty.",
                    roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                )
            elif 'spListName' not in action.inputs or not action.inputs['spListName'].strip():
                botApi.messages.create(
                    text="Sharepoint List name cannot be empty.",
                    roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                )
            elif 'spFolderName' not in action.inputs or not action.inputs['spFolderName'].strip():
                botApi.messages.create(
                    text="Sharepoint Lists folder name cannot be empty.",
                    roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                )
            else:
                try:
                    spSiteURL = action.inputs['spSiteURL'].strip()
                    spListName = action.inputs['spListName'].strip()
                    spFolderName = action.inputs['spFolderName'].strip()

                    # TODO check if any verification of the provided parameters is needed before saving

                    try:
                        saveSharepointParams(spSiteURL, spListName, spFolderName)
                        # send confirmation message
                        botApi.messages.create(
                            markdown=f"New Sharepoint parameters are set:\nSite URL: ``{spSiteURL}``\nList name: ``{spListName}``\nFolder name: ``{spFolderName}``",
                            roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                        )
                        # resend greeting card
                        botApi.messages.create(
                            text=greetingCard.fallbackText,
                            roomId=os.getenv("WEBEX_BOT_ROOM_ID"),
                            attachments=[greetingCard]
                        )
                    except Exception:
                        botApi.messages.create(
                            text="Could not save new Sharepoint parameters to Parameter Store. Check local AWS configuration.",
                            roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                        )
                except Exception:
                    botApi.messages.create(
                        text="That Sharepoint Lists folder name did not work. Try again.",
                        roomId=os.getenv("WEBEX_BOT_ROOM_ID")
                    )

        # "Authorize Webex" action
        if action.type == "submit" and action.inputs['act'] == "authorize webex":
            try:
                # get a fresh Webex Integration access token
                access_token = getWebexIntegrationToken(
                    webex_integration_client_id=os.getenv("WEBEX_INTEGRATION_CLIENT_ID"),
                    webex_integration_client_secret=os.getenv("WEBEX_INTEGRATION_CLIENT_SECRET")
                )

                # get information about the current authorized Webex user
                webexApi = webexteamssdk.WebexTeamsAPI(access_token)
                webexMe = webexApi.people.me()
                webexEmail = webexMe.emails[0]
                webexDisplayName = webexMe.displayName
            except Exception:
                # haven't been authorized yet
                webexEmail = ""
                webexDisplayName = "Not authorized yet"

            if os.getenv("FLASK_ENV") == "development":
                # dev
                authUrl = "http://localhost:5000" + "/auth"
            else:
                # prod
                authUrl = url_for("auth", _external=True)

            card = AdaptiveCard(
                fallbackText="Adaptive cards feature is required to use me.",
                body=[
                    TextBlock(
                        text="Webex integration authorization",
                        weight=FontWeight.BOLDER,
                        size=FontSize.MEDIUM,
                    ),
                    TextBlock(
                        text="Webex integration is used to create Webex Webinar sessions. This is the user currently authorized to create sessions. You can update authorization here.",
                        wrap=True,
                    ),
                    FactSet(
                        facts=[
                            Fact(
                                title="Name",
                                value=webexDisplayName
                            ),
                            Fact(
                                title="Email",
                                value=webexEmail
                            )
                        ]
                    )
                ],
                actions=[
                    ShowCard(
                        title="Change",
                        card=AdaptiveCard(
                            body=[
                                TextBlock(
                                    text="Click the button below and complete authorization process in your browser.",
                                    wrap=True,
                                ),
                            ],
                            actions=[
                                OpenUrl(title="Authorize", url=authUrl),
                                ShowCard(
                                    title="Copy URL",
                                    card=AdaptiveCard(
                                        body=[
                                            TextBlock(
                                                text=authUrl,
                                                wrap=True,
                                            ),
                                        ],
                                    )
                                )
                            ]
                        )
                    )
                ]
            )    # /AdaptiveCard

            botApi.messages.create(text="Could not send the action card", roomId=os.getenv("WEBEX_BOT_ROOM_ID"), attachments=[card])

    return "webhook accepted"
