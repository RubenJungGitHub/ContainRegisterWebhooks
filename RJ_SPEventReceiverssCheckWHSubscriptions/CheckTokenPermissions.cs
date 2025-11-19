namespace RJ_SPEventReceiversASPWebApp

{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;


        public class TokenPermissionChecker
        {
            public static bool HasRequiredPermissions(string accessToken, string[] requiredPermissions)
            {
                var handler = new JwtSecurityTokenHandler();
                var token = handler.ReadJwtToken(accessToken);

                // Permissions for delegated scopes are usually in "scp"
                var scpClaims = token.Claims.Where(c => c.Type == "scp").Select(c => c.Value).FirstOrDefault()?.Split(' ');

                // Permissions for application roles are usually in "roles"
                var roleClaims = token.Claims.Where(c => c.Type == "roles").Select(c => c.Value).ToArray();

                // Combine all permissions found
                var allPermissions = (scpClaims ?? Array.Empty<string>()).Concat(roleClaims ?? Array.Empty<string>()).ToArray();

                // Check if all required permissions are present
                foreach (var requiredPermission in requiredPermissions)
                {
                    if (!allPermissions.Contains(requiredPermission, StringComparer.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Missing permission: {requiredPermission}");
                        return false;
                    }
                }

                return true;
            }
        }
    }
