import QtQuick 2.15
import QtQuick.Controls 2.15

ApplicationWindow {
    visible: true
    width: 800
    height: 600
    title: "Dashboard App"

    // Define a TabView to hold the tabs
    TabView {
        anchors.fill: parent

        // Tab 1
        Tab {
            title: "Tab 1"

            // Content for Tab 1
            Rectangle {
                color: "lightblue"
                width: parent.width
                height: parent.height

                Text {
                    text: "Content for Tab 1"
                    anchors.centerIn: parent
                }
            }
        }

        // Tab 2
        Tab {
            title: "Tab 2"

            // Content for Tab 2
            Rectangle {
                color: "lightgreen"
                width: parent.width
                height: parent.height

                Text {
                    text: "Content for Tab 2"
                    anchors.centerIn: parent
                }
            }
        }

        // Tab 3
        Tab {
            title: "Tab 3"

            // Content for Tab 3
            Rectangle {
                color: "lightcoral"
                width: parent.width
                height: parent.height

                Text {
                    text: "Content for Tab 3"
                    anchors.centerIn: parent
                }
            }
        }
    }
}
import QtQuick 2.15
import QtQuick.Controls 2.15

ApplicationWindow {
    visible: true
    width: 800
    height: 600
    title: "Dashboard App"

    // Define a TabView to hold the tabs
    TabView {
        anchors.fill: parent

        // Tab 1
        Tab {
            title: "Tab 1"

            // Content for Tab 1
            Rectangle {
                color: "lightblue"
                width: parent.width
                height: parent.height

                Text {
                    text: "Content for Tab 1"
                    anchors.centerIn: parent
                }
            }
        }

        // Tab 2
        Tab {
            title: "Tab 2"

            // Content for Tab 2
            Rectangle {
                color: "lightgreen"
                width: parent.width
                height: parent.height

                Text {
                    text: "Content for Tab 2"
                    anchors.centerIn: parent
                }
            }
        }

        // Tab 3
        Tab {
            title: "Tab 3"

            // Content for Tab 3
            Rectangle {
                color: "lightcoral"
                width: parent.width
                height: parent.height

                Text {
                    text: "Content for Tab 3"
                    anchors.centerIn: parent
                }
            }
        }
    }
}
