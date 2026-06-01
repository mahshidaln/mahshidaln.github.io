# Deployment Notes

This is a static GitHub Pages site.

## Preserve Music-STAR

The Music-STAR project is still hosted at:

```text
Music-STAR.html
```

Do not rename this file unless you also add a redirect from the old URL.

## Custom Domain

After choosing the final domain, add a root-level `CNAME` file containing only the domain name:

```text
example.com
```

Then configure DNS at the domain registrar:

- For an apex domain, add GitHub Pages `A` records.
- For a `www` subdomain, add a `CNAME` record pointing to `mahshidaln.github.io`.

The domain can then be confirmed in the GitHub Pages settings for this repository.
