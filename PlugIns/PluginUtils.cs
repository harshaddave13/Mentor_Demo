
using Microsoft.Xrm.Sdk;
using System;
using System.Linq.Expressions;

public static class PluginUtils
{
    /// <summary>
    /// Merge the Target with an Pre-Image to get the updated values and nonupdated values in one object
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="image"></param>
    /// <returns></returns>
    public static T MergeTargetAndImage<T>(Entity target, Entity image) where T : Entity
    {
        T mergedEntity = null;
        mergedEntity = new Entity(target.LogicalName).ToEntity<T>();
        mergedEntity.Id = target.Id;

        if (image != null)
        {
            foreach (string key in image.Attributes.Keys)
            {
                mergedEntity[key] = image[key];
            }

        }

        foreach (string attribute in target.Attributes.Keys)
        {
            if (mergedEntity.Attributes.Contains(attribute))
            {
                mergedEntity.Attributes[attribute] = target.Attributes[attribute];
            }
            else
            {
                mergedEntity.Attributes.Add(attribute, target.Attributes[attribute]);
            }
        }


        return mergedEntity;


    }

    public static string GetAttributeName<T>(Expression<Func<T>> propertyExpression)
    {
        var body = propertyExpression.Body as MemberExpression;
        var expression = body.Expression as ConstantExpression;
        return body.Member.Name.ToLowerInvariant();
    }
}

