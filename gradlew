#!/bin/sh
DIR="$(cd "$(dirname "$0")" && pwd)"
exec java -Dorg.gradle.appname=gradlew -classpath "$DIR/gradle/wrapper/gradle-wrapper.jar" org.gradle.wrapper.GradleWrapperMain "$@"
