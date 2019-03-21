# BugUpdate
Real-time status updates on a customer's bugs.

Each customer has a "Z channel" in the TAP100 team, in which an incoming webhook is installed. The MSTeams Azure DevOps has service hooks that notify this Azure Function when a bug is created or resolved. This function determines if the bug belongs to any of the customers, then sends a MessageCard about the bug and its status to the right customer's webhook.
