from constants import NSMAP


def get_first_sub_element(element, xpath):
    if element is None:
        return None

    sub_elements = element.xpath(xpath, namespaces=NSMAP)
    if not len(sub_elements):
        return None
    return sub_elements[0]
