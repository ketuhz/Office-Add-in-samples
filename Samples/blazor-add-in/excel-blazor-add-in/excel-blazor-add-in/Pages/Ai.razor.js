/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
export function scrollToBottom(element) {
    if (element) {
        element.scrollTop = element.scrollHeight;
    }
}
