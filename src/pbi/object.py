import secrets


class _PbiObject:
    name_prefix = ''

    def _generate_name(self):
        """
        Generates a random name using the prefix and a hexadecimal
        :return: a string
        """
        return self.name_prefix + secrets.token_hex(10)
