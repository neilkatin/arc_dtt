#!/usr/bin/env python


import argparse
import logging
import os
from concurrent import futures
from typing import Callable

from google.cloud import pubsub_v1

import neil_tools.init_config
import neil_tools.init_logging

import config as config_static


# these should probably be in a config file, but for now they are fine here...
PROJECT_ID = 'arc-transportation-reports'
DRO_TOPIC_ID = 'process-dro'
DRO_SUBSCRIPTION_ID = 'process-dro-sub'
SUB_TIMEOUT = 5.0

def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    if args.pub:
        publish()
    elif args.sub:
        subscribe()
    else:
        log.fatal("Missing required argument: --pub or --sub")
        return


def publish():
    log.debug("Publish...")

    publisher = pubsub_v1.PublisherClient()
    topic_path = publisher.topic_path(PROJECT_ID, DRO_TOPIC_ID)
    publish_futures = []

    def get_callback(
        publish_future: pubsub_v1.futures.Future, data: str
    ) -> Callable[[pubsub_v1.publisher.futures.Future], None]:
        def callback(publish_future: pubsub_v1.publisher.futures.Future) -> None:
            try:
                # Wait 60 seconds for the publish call to succeed.
                print(publish_future.result(timeout=60))
            except futures.TimeoutError:
                print(f"Publishing {data} timed out.")

        return callback

    for i in range(3):
        data = str(i)
        # When you publish a message, the client returns a future.
        publish_future = publisher.publish(topic_path, data.encode("utf-8"))

        # Non-blocking. Publish failures are handled in the callback function.
        publish_future.add_done_callback(get_callback(publish_future, data))
        publish_futures.append(publish_future)

    # Wait for all the publish futures to resolve before exiting.
    futures.wait(publish_futures, return_when=futures.ALL_COMPLETED)

    log.info(f"Published messages with error handler to {topic_path}.")


def subscribe():
    log.debug("Subscribe...")

    subscriber = pubsub_v1.SubscriberClient()
    # The `subscription_path` method creates a fully qualified identifier
    # in the form `projects/{project_id}/subscriptions/{subscription_id}`
    subscription_path = subscriber.subscription_path(PROJECT_ID, DRO_SUBSCRIPTION_ID)

    log.debug(f"subscription_path '{ subscription_path }'")

    def callback(message: pubsub_v1.subscriber.message.Message) -> None:
        print(f"Received {message}.")
        message.ack()

    streaming_pull_future = subscriber.subscribe(subscription_path, callback=callback)
    print(f"Listening for messages on {subscription_path}..\n")

    # Wrap subscriber in a 'with' block to automatically call close() when done.
    with subscriber:
        try:
            # When `timeout` is not set, result() will block indefinitely,
            # unless an exception is encountered first.
            streaming_pull_future.result(timeout=SUB_TIMEOUT)
        except futures.TimeoutError:
            streaming_pull_future.cancel()  # Trigger the shutdown.
            streaming_pull_future.result()  # Block until the shutdown is complete.




def parse_args():
    parser = argparse.ArgumentParser(
            description="tools to support Disaster Transportation Tools reporting",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--pub", help="Publish a message", action="store_true")
    group.add_argument("--sub", help="Subscribe to a message", action="store_true")

    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--save-input", help="Save a copy of server inputs", action="store_true")
    group.add_argument("--cached-input", help="Use cached server input", action="store_true")

    args = parser.parse_args()

    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)

